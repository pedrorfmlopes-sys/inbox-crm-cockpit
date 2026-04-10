import type { GroupStorageLegacyProvider, GroupStorageLocationPointer, GroupStorageMode, GroupStorageSettings } from "../types";

export type GroupStorageProviderAdapter = {
  id: GroupStorageMode;
  legacyProvider: GroupStorageLegacyProvider;
  describePrimary(settings: GroupStorageSettings): GroupStorageLocationPointer;
  describeRemote(settings: GroupStorageSettings): GroupStorageLocationPointer | null;
};
