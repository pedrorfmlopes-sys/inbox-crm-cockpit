import type { GroupStorageProviderAdapter } from "./providerTypes";

export const chosenFolderProvider: GroupStorageProviderAdapter = {
  id: "chosen_folder",
  legacyProvider: "local",
  describePrimary(settings) {
    const basePath = String(settings.chosenFolder.path || settings.baseFolderPath || "").trim();
    const isLibrary = settings.chosenFolder.kind === "document_library";
    return {
      kind: isLibrary ? "document_library" : "filesystem",
      provider: isLibrary ? "onedrive" : "local",
      label: isLibrary ? "Pasta escolhida (biblioteca/document library)" : "Pasta escolhida",
      basePath: basePath || undefined,
      isRemote: isLibrary,
      isConfigured: Boolean(basePath),
    };
  },
  describeRemote() {
    return null;
  },
};
