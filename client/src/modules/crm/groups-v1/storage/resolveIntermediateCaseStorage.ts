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
      reason: "Storage intermedio desligado nos Settings da aba Groups.",
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
      reason: "Sem namespace configurado para o IndexedDB local; o caso intermedio fica apenas em memoria nesta fase.",
    };
  }

  const adapter = createIndexedDbIntermediateCaseStorageAdapter({ namespace });
  return {
    adapter,
    repository: createIntermediateCaseRepository(adapter),
    mode: "indexeddb",
    availability: "ready",
    namespace,
    reason: "Storage intermedio real desta fase via IndexedDB local do host, namespaced pela chave configurada.",
  };
}
