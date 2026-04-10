import type { GroupStorageProviderAdapter } from "./providerTypes";

export const localDeviceProvider: GroupStorageProviderAdapter = {
  id: "local_device",
  legacyProvider: "local",
  describePrimary(settings) {
    const basePath = String(settings.localDevice.rootPath || settings.baseFolderPath || "").trim();
    return {
      kind: "local_device",
      provider: "local",
      label: "Local neste PC",
      basePath: basePath || undefined,
      isRemote: false,
      isConfigured: Boolean(basePath),
    };
  },
  describeRemote() {
    return null;
  },
};
