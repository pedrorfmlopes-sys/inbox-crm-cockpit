import { chosenFolderProvider } from "./chosenFolderProvider";
import { localDeviceProvider } from "./localDeviceProvider";
import type { GroupStorageProviderAdapter } from "./providerTypes";

export const hybridProvider: GroupStorageProviderAdapter = {
  id: "hybrid",
  legacyProvider: "local",
  describePrimary(settings) {
    return settings.hybrid.primaryTarget === "chosen_folder"
      ? chosenFolderProvider.describePrimary(settings)
      : localDeviceProvider.describePrimary(settings);
  },
  describeRemote() {
    return {
      kind: "supabase",
      provider: "supabase",
      label: "Supabase (promocao remota)",
      isRemote: true,
      isConfigured: true,
    };
  },
};
