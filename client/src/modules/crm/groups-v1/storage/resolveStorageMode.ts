import { resolveGroupAttachmentStoragePolicy } from "./attachmentPolicy";
import { chosenFolderProvider } from "./providers/chosenFolderProvider";
import { hybridProvider } from "./providers/hybridProvider";
import { localDeviceProvider } from "./providers/localDeviceProvider";
import type { GroupStorageProviderAdapter } from "./providers/providerTypes";
import { supabaseProvider } from "./providers/supabaseProvider";
import { resolveGroupPromotionPolicy } from "./promotionPolicy";
import { GROUP_STORAGE_MODE_LABELS } from "./modes";
import { normalizeGroupStorageSettings } from "./settings";
import type { GroupStorageLocationPointer, GroupStorageSettings } from "./types";

function pickProvider(settings: GroupStorageSettings): GroupStorageProviderAdapter {
  if (settings.mode === "local_device") return localDeviceProvider;
  if (settings.mode === "chosen_folder") return chosenFolderProvider;
  if (settings.mode === "hybrid") return hybridProvider;
  return supabaseProvider;
}

export type ResolvedGroupStorageRuntime = {
  settings: GroupStorageSettings;
  mode: GroupStorageSettings["mode"];
  modeLabel: string;
  primaryLocation: GroupStorageLocationPointer;
  remotePromotionLocation: GroupStorageLocationPointer | null;
  attachmentPolicy: ReturnType<typeof resolveGroupAttachmentStoragePolicy>;
  promotionPolicy: ReturnType<typeof resolveGroupPromotionPolicy>;
  legacyBridge: {
    provider: GroupStorageSettings["provider"];
    baseFolderPath: string;
    usesFileBackedStorage: boolean;
  };
};

export function resolveGroupStorageRuntime(settingsLike?: { groupStorage?: Partial<GroupStorageSettings> | null } | Partial<GroupStorageSettings> | null): ResolvedGroupStorageRuntime {
  const raw = settingsLike && typeof settingsLike === "object" && "groupStorage" in settingsLike
    ? settingsLike.groupStorage || null
    : settingsLike;
  const settings = normalizeGroupStorageSettings(raw || null);
  const provider = pickProvider(settings);
  const primaryLocation = provider.describePrimary(settings);
  const remotePromotionLocation = provider.describeRemote(settings);
  const attachmentPolicy = resolveGroupAttachmentStoragePolicy(settings);
  const promotionPolicy = resolveGroupPromotionPolicy(settings, primaryLocation, remotePromotionLocation);
  return {
    settings,
    mode: settings.mode,
    modeLabel: GROUP_STORAGE_MODE_LABELS[settings.mode],
    primaryLocation,
    remotePromotionLocation,
    attachmentPolicy,
    promotionPolicy,
    legacyBridge: {
      provider:
        primaryLocation.provider === "local" || primaryLocation.provider === "onedrive" || primaryLocation.provider === "cloud"
          ? primaryLocation.provider
          : provider.legacyProvider,
      baseFolderPath: String(primaryLocation.basePath || "").trim(),
      usesFileBackedStorage: primaryLocation.provider === "local" || primaryLocation.provider === "onedrive",
    },
  };
}

export function getGroupAttachmentStorageOptions(settingsLike?: { groupStorage?: Partial<GroupStorageSettings> | null } | Partial<GroupStorageSettings> | null): {
  attachmentStorageProvider: GroupStorageSettings["provider"];
  attachmentStorageBasePath: string;
} {
  const runtime = resolveGroupStorageRuntime(settingsLike);
  return {
    attachmentStorageProvider: runtime.legacyBridge.provider,
    attachmentStorageBasePath: runtime.legacyBridge.baseFolderPath,
  };
}
