import { resolveGroupAttachmentStoragePolicy } from "./attachmentPolicy";
import { chosenFolderProvider } from "./providers/chosenFolderProvider";
import { hybridProvider } from "./providers/hybridProvider";
import { localDeviceProvider } from "./providers/localDeviceProvider";
import type { GroupStorageProviderAdapter } from "./providers/providerTypes";
import { supabaseProvider } from "./providers/supabaseProvider";
import { resolveGroupPromotionPolicy } from "./promotionPolicy";
import { GROUP_STORAGE_MODE_LABELS } from "./modes";
import { getGroupsModuleSettings } from "../settings/groupsModuleSettings";
import { isGraphAdminBlockedGroupStorageConfig, normalizeGroupStorageSettings } from "./settings";
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
  projectSupport: {
    supported: boolean;
    requiresGraphOrAdmin: boolean;
    blockingReason: string | null;
  };
  legacyBridge: {
    provider: GroupStorageSettings["provider"];
    baseFolderPath: string;
    usesFileBackedStorage: boolean;
  };
};

function resolveStorageInput(
  settingsLike?:
    | { groups?: { storage?: Partial<GroupStorageSettings> | null } | null }
    | Partial<GroupStorageSettings>
    | null
): Partial<GroupStorageSettings> | null {
  if (!settingsLike || typeof settingsLike !== "object") return settingsLike || null;
  if ("groups" in settingsLike) {
    return getGroupsModuleSettings(settingsLike).storage;
  }
  return settingsLike;
}

export function resolveGroupStorageRuntime(settingsLike?: { groups?: { storage?: Partial<GroupStorageSettings> | null } | null } | Partial<GroupStorageSettings> | null): ResolvedGroupStorageRuntime {
  const settings = normalizeGroupStorageSettings(resolveStorageInput(settingsLike) || null);
  const provider = pickProvider(settings);
  const primaryLocation = provider.describePrimary(settings);
  const remotePromotionLocation = provider.describeRemote(settings);
  const attachmentPolicy = resolveGroupAttachmentStoragePolicy(settings);
  const promotionPolicy = resolveGroupPromotionPolicy(settings, primaryLocation, remotePromotionLocation);
  const requiresGraphOrAdmin = isGraphAdminBlockedGroupStorageConfig(settings);
  const supported = !requiresGraphOrAdmin;
  const blockingReason = requiresGraphOrAdmin
    ? "URL web de OneDrive/SharePoint exige Graph/SharePoint API e permissoes fora do perimetro desta fase."
    : null;
  const legacyProvider = !supported
    ? "cloud"
    : primaryLocation.provider === "local" || primaryLocation.provider === "onedrive" || primaryLocation.provider === "cloud"
      ? primaryLocation.provider
      : provider.legacyProvider;
  const legacyBaseFolderPath = !supported ? "" : String(primaryLocation.basePath || "").trim();
  return {
    settings,
    mode: settings.mode,
    modeLabel: GROUP_STORAGE_MODE_LABELS[settings.mode],
    primaryLocation,
    remotePromotionLocation,
    attachmentPolicy,
    promotionPolicy,
    projectSupport: {
      supported,
      requiresGraphOrAdmin,
      blockingReason,
    },
    legacyBridge: {
      provider: legacyProvider,
      baseFolderPath: legacyBaseFolderPath,
      usesFileBackedStorage: supported && (primaryLocation.provider === "local" || primaryLocation.provider === "onedrive"),
    },
  };
}

export function getGroupAttachmentStorageOptions(settingsLike?: { groups?: { storage?: Partial<GroupStorageSettings> | null } | null } | Partial<GroupStorageSettings> | null): {
  attachmentStorageProvider: GroupStorageSettings["provider"];
  attachmentStorageBasePath: string;
} {
  const runtime = resolveGroupStorageRuntime(settingsLike);
  return {
    attachmentStorageProvider: runtime.legacyBridge.provider,
    attachmentStorageBasePath: runtime.legacyBridge.baseFolderPath,
  };
}
