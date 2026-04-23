import { createIndexedDbIntermediateCaseStorageAdapter } from "./intermediateCaseIndexedDbAdapter";
import {
  createInMemoryIntermediateCaseStorageAdapter,
  createIntermediateCaseRepository,
  type IntermediateCaseStorageAdapter,
} from "./intermediateCaseRepository";
import { createServerBackedIntermediateCaseStorageAdapter } from "./intermediateCaseServerAdapter";
import type { GroupsTabSettings } from "../settings/groupsTabSettings";

export const GROUPS_ADDIN_LOCAL_INTERMEDIATE_NAMESPACE = "groups_v1_addin_local";

export type ResolvedIntermediateCaseStorage = {
  adapter: IntermediateCaseStorageAdapter;
  repository: ReturnType<typeof createIntermediateCaseRepository>;
  mode: "filesystem" | "indexeddb" | "memory";
  availability: "ready" | "fallback_memory" | "disabled";
  locationPath?: string;
  reason: string;
};

function canUseIndexedDbStorage(): boolean {
  try {
    return typeof globalThis.indexedDB !== "undefined" && globalThis.indexedDB !== null;
  } catch {
    return false;
  }
}

export function resolveIntermediateCaseStorage(settings: GroupsTabSettings): ResolvedIntermediateCaseStorage {
  if (settings.storageMode === "disabled") {
    const adapter = createInMemoryIntermediateCaseStorageAdapter();
    return {
      adapter,
      repository: createIntermediateCaseRepository(adapter),
      mode: "memory",
      availability: "disabled",
      reason: "Storage intermédio desligado nos Settings da aba Groups.",
    };
  }

  const locationPath = String(settings.baseFolderPath || "").trim();
  if (locationPath) {
    const adapter = createServerBackedIntermediateCaseStorageAdapter({ basePath: locationPath });
    return {
      adapter,
      repository: createIntermediateCaseRepository(adapter),
      mode: "filesystem",
      availability: "ready",
      locationPath,
      reason: "Pasta local definida: o caso intermédio grava logo nessa localização e a reabertura lê daí.",
    };
  }

  if (canUseIndexedDbStorage()) {
    const adapter = createIndexedDbIntermediateCaseStorageAdapter({
      namespace: GROUPS_ADDIN_LOCAL_INTERMEDIATE_NAMESPACE,
    });
    return {
      adapter,
      repository: createIntermediateCaseRepository(adapter),
      mode: "indexeddb",
      availability: "ready",
      reason: "Sem pasta intermédia definida: o fallback principal passa a ser o storage local do add-in em IndexedDB.",
    };
  }

  const adapter = createInMemoryIntermediateCaseStorageAdapter();
  return {
    adapter,
    repository: createIntermediateCaseRepository(adapter),
    mode: "memory",
    availability: "fallback_memory",
    reason: "Nem pasta local nem IndexedDB do add-in estão disponíveis; o caso intermédio caiu para memória apenas como fallback técnico.",
  };
}
