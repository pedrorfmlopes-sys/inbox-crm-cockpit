import type { GroupStorageLegacyProvider, GroupStorageMode, GroupStorageSettings } from "./types";

export type { GroupStorageLegacyProvider, GroupStorageMode, GroupStorageSettings } from "./types";

export const DEFAULT_GROUP_STORAGE_SETTINGS: GroupStorageSettings = {
  mode: "supabase",
  provider: "cloud",
  baseFolderPath: "",
  autoCreateFolderOnGroupCreate: true,
  ignoreInlineAttachments: true,
  suggestedViewer: "inline",
  attachmentPromptThresholdMb: 10,
  localDevice: {
    rootPath: "",
  },
  chosenFolder: {
    path: "",
    kind: "filesystem",
  },
  supabase: {
    allowPromotion: true,
    promoteManifestOnSave: true,
    promoteAttachmentMetadataOnSave: false,
    promoteAttachmentBinaryOnSave: false,
  },
  hybrid: {
    primaryTarget: "chosen_folder",
    promoteManifestOnSave: true,
    promoteAttachmentMetadataOnSave: false,
  },
};

export function normalizeGroupStorageLegacyProvider(value: unknown): GroupStorageLegacyProvider {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "local" || normalized === "onedrive") return normalized;
  return "cloud";
}

export function normalizeGroupStorageMode(value: unknown): GroupStorageMode | null {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "local_device" || normalized === "chosen_folder" || normalized === "hybrid" || normalized === "supabase") {
    return normalized;
  }
  return null;
}

function detectChosenFolderKind(pathValue: string, explicitKind: unknown): GroupStorageSettings["chosenFolder"]["kind"] {
  const normalizedKind = String(explicitKind || "").trim().toLowerCase();
  if (normalizedKind === "document_library") return "document_library";
  if (normalizedKind === "filesystem") return "filesystem";
  return "filesystem";
}

export function isGraphAdminBlockedGroupStorageConfig(
  input: Partial<GroupStorageSettings> | null | undefined
): boolean {
  const provider = normalizeGroupStorageLegacyProvider(input?.provider);
  const mode = normalizeGroupStorageMode(input?.mode)
    || inferModeFromLegacy(provider, String(input?.baseFolderPath || "").trim());
  if (mode === "supabase" || mode === "local_device") {
    return false;
  }
  if (mode === "hybrid" && input?.hybrid?.primaryTarget === "local_device") {
    return false;
  }
  const chosenPath = String(input?.chosenFolder?.path || input?.baseFolderPath || "").trim();
  const explicitKind = String(input?.chosenFolder?.kind || "").trim().toLowerCase();
  return explicitKind === "document_library" || /^https?:\/\//i.test(chosenPath);
}

function inferModeFromLegacy(provider: GroupStorageLegacyProvider, baseFolderPath: string): GroupStorageMode {
  if (provider === "onedrive") return "chosen_folder";
  if (provider === "local") return baseFolderPath ? "chosen_folder" : "local_device";
  return "supabase";
}

function clampThresholdMb(value: unknown): number {
  const numeric = Number(value || 0);
  if (!Number.isFinite(numeric) || numeric <= 0) return DEFAULT_GROUP_STORAGE_SETTINGS.attachmentPromptThresholdMb;
  return Math.min(Math.max(Math.round(numeric), 1), 250);
}

export function normalizeGroupStorageSettings(input: Partial<GroupStorageSettings> | null | undefined): GroupStorageSettings {
  const baseFolderPath = String(input?.baseFolderPath || "").trim();
  const localRoot = String(input?.localDevice?.rootPath || "").trim();
  const chosenPath = String(input?.chosenFolder?.path || baseFolderPath || "").trim();
  const provider = normalizeGroupStorageLegacyProvider(input?.provider);
  const mode = normalizeGroupStorageMode(input?.mode) || inferModeFromLegacy(provider, baseFolderPath);
  const chosenKind = detectChosenFolderKind(chosenPath, input?.chosenFolder?.kind);

  const settings: GroupStorageSettings = {
    mode,
    provider,
    baseFolderPath,
    autoCreateFolderOnGroupCreate: input?.autoCreateFolderOnGroupCreate !== false,
    ignoreInlineAttachments: input?.ignoreInlineAttachments !== false,
    suggestedViewer: input?.suggestedViewer === "system" ? "system" : "inline",
    attachmentPromptThresholdMb: clampThresholdMb(input?.attachmentPromptThresholdMb),
    localDevice: {
      rootPath: localRoot,
    },
    chosenFolder: {
      path: chosenPath,
      kind: chosenKind,
    },
    supabase: {
      allowPromotion: input?.supabase?.allowPromotion !== false,
      promoteManifestOnSave: input?.supabase?.promoteManifestOnSave !== false,
      promoteAttachmentMetadataOnSave: input?.supabase?.promoteAttachmentMetadataOnSave === true,
      promoteAttachmentBinaryOnSave: input?.supabase?.promoteAttachmentBinaryOnSave === true,
    },
    hybrid: {
      primaryTarget: input?.hybrid?.primaryTarget === "local_device" ? "local_device" : "chosen_folder",
      promoteManifestOnSave: input?.hybrid?.promoteManifestOnSave !== false,
      promoteAttachmentMetadataOnSave: input?.hybrid?.promoteAttachmentMetadataOnSave === true,
    },
  };

  if (settings.mode === "supabase") {
    settings.provider = "cloud";
    settings.baseFolderPath = "";
    return settings;
  }

  if (settings.mode === "local_device") {
    settings.provider = "local";
    settings.baseFolderPath = settings.localDevice.rootPath;
    settings.hybrid.primaryTarget = "local_device";
    return settings;
  }

  if (settings.mode === "chosen_folder") {
    settings.provider = settings.chosenFolder.kind === "document_library" ? "onedrive" : "local";
    settings.baseFolderPath = settings.chosenFolder.path;
    return settings;
  }

  settings.provider = settings.hybrid.primaryTarget === "local_device"
    ? "local"
    : settings.chosenFolder.kind === "document_library"
      ? "onedrive"
      : "local";
  settings.baseFolderPath = settings.hybrid.primaryTarget === "local_device"
    ? settings.localDevice.rootPath
    : settings.chosenFolder.path;
  return settings;
}
