import type { GroupStorageProviderAdapter } from "./providerTypes";

export const supabaseProvider: GroupStorageProviderAdapter = {
  id: "supabase",
  legacyProvider: "cloud",
  describePrimary() {
    return {
      kind: "supabase",
      provider: "supabase",
      label: "Supabase",
      isRemote: true,
      isConfigured: true,
    };
  },
  describeRemote() {
    return null;
  },
};
