import { createIndexedDbIntermediateCaseStorageAdapter } from "./intermediateCaseIndexedDbAdapter";
import { createInMemoryIntermediateCaseStorageAdapter, createIntermediateCaseRepository, type IntermediateCaseStorageAdapter } from "./intermediateCaseRepository";
import type { GroupsTabSettings } from "../settings/groupsTabSettings";

export type ResolvedIntermediateCaseStorage = {
  adapter: IntermediateCaseStorageAdapter;
  repository: ReturnType<typeof createIntermediateCaseRepository>;
  mode: "indexeddb" | "memory";
  availability: "ready" | "missing_location" | "disabled";
  namespace?: string;
  reason: string;
};

export function resolveIntermediateCaseStorage(settings: GroupsTabSettings): ResolvedIntermediateCaseStorage {
  if (settings.storageMode === "disabled") {
    const adapter = createInMemoryIntermediateCaseStorageAdapter();
    return {
      adapter,
      repository: createIntermediateCaseRepository(adapter),
      mode: "memory",
      availability: "disabled",
      reason: "Storage intermédio desligado nos Settings.",
    };
  }

  const namespace = String(settings.baseFolderPath || "").trim();
  if (!namespace) {
    const adapter = createInMemoryIntermediateCaseStorageAdapter();
    return {
      adapter,
      repository: createIntermediateCaseRepository(adapter),
      mode: "memory",
      availability: "missing_location",
      reason: "Sem bridge real para a localização escolhida; fallback transitório em memória até existir integração de pasta/OneDrive.",
    };
  }

  const adapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace });
  return {
    adapter,
    repository: createIntermediateCaseRepository(adapter),
    mode: "indexeddb",
    availability: "ready",
    namespace,
    reason: "Storage intermédio real suportado nesta ronda via IndexedDB do host, namespaced pela localização configurada.",
  };
}
