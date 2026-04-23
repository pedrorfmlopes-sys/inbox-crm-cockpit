import type { IntermediateCaseBinaryStorageAdapter } from "./intermediateCaseRepository";

const INTERMEDIATE_CASE_API_TIMEOUT_MS = 30000;

type IntermediateStorageRequestPayload = {
  basePath: string;
  path?: string;
  prefix?: string;
  content?: string;
  contentBase64?: string;
  contentType?: string;
  initialPath?: string;
  description?: string;
};

export type IntermediateStorageFolderPickerResult = {
  supported: boolean;
  selected: boolean;
  cancelled?: boolean;
  path: string;
  normalizedPath: string;
  picker?: string | null;
  reason?: string;
  validation?: {
    supported?: boolean;
    blockingReason?: string;
    normalizedBasePath?: string;
    notes?: string[];
  } | null;
};

function normalizeText(value: unknown): string {
  return String(value || "").trim();
}

async function requestIntermediateStorageJson<T>(
  route: string,
  payload: IntermediateStorageRequestPayload
): Promise<T> {
  const controller = new AbortController();
  const timeoutMessage = `Pedido de storage intermédio excedeu ${Math.round(INTERMEDIATE_CASE_API_TIMEOUT_MS / 1000)}s: ${route}`;
  const id = setTimeout(() => controller.abort(timeoutMessage), INTERMEDIATE_CASE_API_TIMEOUT_MS);

  try {
    const res = await fetch(route, {
      method: "POST",
      signal: controller.signal,
      headers: {
        "Content-Type": "application/json",
      },
      body: JSON.stringify(payload),
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

function decodeBase64ToBlob(base64: string, contentType?: string): Blob | null {
  const raw = normalizeText(base64).replace(/^data:[^,]+,/, "");
  if (!raw) return null;
  try {
    const binary = atob(raw);
    const bytes = new Uint8Array(binary.length);
    for (let index = 0; index < binary.length; index += 1) {
      bytes[index] = binary.charCodeAt(index);
    }
    return new Blob([bytes], { type: normalizeText(contentType) || "application/octet-stream" });
  } catch {
    return null;
  }
}

async function encodeBinaryToBase64(content: Blob | Uint8Array | ArrayBuffer): Promise<string> {
  const bytes = content instanceof Blob
    ? new Uint8Array(await content.arrayBuffer())
    : content instanceof ArrayBuffer
      ? new Uint8Array(content)
      : content;
  let binary = "";
  for (let index = 0; index < bytes.length; index += 1) {
    binary += String.fromCharCode(bytes[index]);
  }
  return btoa(binary);
}

export function createServerBackedIntermediateCaseStorageAdapter(input: {
  basePath: string;
}): IntermediateCaseBinaryStorageAdapter {
  const basePath = normalizeText(input.basePath);

  async function post<T>(route: string, payload: Omit<IntermediateStorageRequestPayload, "basePath">): Promise<T> {
    return await requestIntermediateStorageJson<T>(route, {
      basePath,
      ...payload,
    });
  }

  return {
    async readText(path) {
      const response = await post<{ content?: string | null }>("/api/links/groups/intermediate-storage/read-text", { path });
      return typeof response?.content === "string" ? response.content : null;
    },
    async writeText(path, content) {
      await post("/api/links/groups/intermediate-storage/write-text", {
        path,
        content,
      });
    },
    async readBinary(path) {
      const response = await post<{ contentBase64?: string | null; contentType?: string | null }>(
        "/api/links/groups/intermediate-storage/read-binary",
        { path }
      );
      return response?.contentBase64
        ? decodeBase64ToBlob(response.contentBase64, response.contentType || undefined)
        : null;
    },
    async writeBinary(path, content) {
      await post("/api/links/groups/intermediate-storage/write-binary", {
        path,
        contentBase64: await encodeBinaryToBase64(content),
        contentType: content instanceof Blob ? content.type : undefined,
      });
    },
    async deleteTree(path) {
      await post("/api/links/groups/intermediate-storage/delete-tree", { path });
    },
    async listPaths(prefix) {
      const response = await post<{ paths?: string[] }>("/api/links/groups/intermediate-storage/list-paths", { prefix });
      return Array.isArray(response?.paths) ? response.paths : [];
    },
  };
}

export async function pickIntermediateCaseStorageFolder(input?: {
  initialPath?: string;
  description?: string;
}): Promise<IntermediateStorageFolderPickerResult> {
  const response = await requestIntermediateStorageJson<{ result?: IntermediateStorageFolderPickerResult }>(
    "/api/links/groups/intermediate-storage/pick-folder",
    {
      basePath: "",
      initialPath: normalizeText(input?.initialPath),
      description: normalizeText(input?.description),
    }
  );
  return response?.result || {
    supported: false,
    selected: false,
    path: "",
    normalizedPath: "",
    picker: null,
    reason: "O picker de pasta nao devolveu resultado.",
    validation: null,
  };
}
