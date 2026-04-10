export type GroupStorageMode = "supabase" | "local_device" | "chosen_folder" | "hybrid";
export type GroupStorageLegacyProvider = "cloud" | "local" | "onedrive";
export type GroupStorageLocationKind =
  | "session"
  | "supabase"
  | "local_device"
  | "filesystem"
  | "document_library"
  | "hybrid";
export type GroupAttachmentBinaryStrategy = "embed_binary" | "store_reference" | "prompt_user" | "skip_inline";
export type GroupPromotionState = "not_requested" | "pending" | "promoted" | "partial" | "skipped" | "blocked";
export type GroupPromotionScope = "manifest" | "email_metadata" | "attachment_metadata" | "attachment_binary";
export type GroupWorksetAttachmentSelection = "selected" | "rejected" | "pending";

export type GroupWorksetFilterSnapshot = {
  query?: string;
  fromEmail?: string;
  labels?: string[];
  dateFromIso?: string;
  dateToIso?: string;
  attachmentMode?: "all" | "with" | "without";
  groupMode?: "all" | "with_group" | "without_group";
};

export type GroupPreparedAttachmentDescriptor = {
  key: string;
  emailKey: string;
  name: string;
  contentType?: string;
  size?: number;
  isInline?: boolean;
  hasContent?: boolean;
  selection: GroupWorksetAttachmentSelection;
  storageDisposition?: "binary" | "reference" | "skip";
  requiresDecision?: boolean;
};

export type GroupStorageLocationPointer = {
  kind: GroupStorageLocationKind;
  provider: GroupStorageLegacyProvider | "supabase" | "hybrid";
  label: string;
  basePath?: string;
  relativePath?: string;
  folderHint?: string;
  isRemote: boolean;
  isConfigured: boolean;
};

export type GroupSupabasePromotionSettings = {
  allowPromotion: boolean;
  promoteManifestOnSave: boolean;
  promoteAttachmentMetadataOnSave: boolean;
  promoteAttachmentBinaryOnSave: boolean;
};

export type GroupHybridStorageSettings = {
  primaryTarget: "local_device" | "chosen_folder";
  promoteManifestOnSave: boolean;
  promoteAttachmentMetadataOnSave: boolean;
};

export type GroupChosenFolderSettings = {
  path: string;
  kind: "filesystem" | "document_library";
};

export type GroupLocalDeviceSettings = {
  rootPath: string;
};

export type GroupStorageSettings = {
  mode: GroupStorageMode;
  provider: GroupStorageLegacyProvider;
  baseFolderPath: string;
  autoCreateFolderOnGroupCreate: boolean;
  ignoreInlineAttachments: boolean;
  suggestedViewer: "system" | "inline";
  attachmentPromptThresholdMb: number;
  localDevice: GroupLocalDeviceSettings;
  chosenFolder: GroupChosenFolderSettings;
  supabase: GroupSupabasePromotionSettings;
  hybrid: GroupHybridStorageSettings;
};

export type GroupStorageSessionDraft = {
  kind: "groups_v1_storage_session_draft";
  version: 1;
  savedAtIso: string;
  storageMode: GroupStorageMode;
  anchorEmailKey: string;
  selectedEmailKeys: string[];
  expandedEmailKeys: string[];
  workingGroupId?: string;
  workingGroupName?: string;
  filters: GroupWorksetFilterSnapshot;
  preparedAttachments: GroupPreparedAttachmentDescriptor[];
};

export type GroupWorksetPromotionStatus = {
  state: GroupPromotionState;
  lastAttemptAtIso?: string;
  promotedScopes: GroupPromotionScope[];
  blockedScopes: GroupPromotionScope[];
  note?: string;
};

export type GroupWorksetManifest = {
  kind: "groups_v1_workset_manifest";
  version: 1;
  worksetKey: string;
  createdAtIso: string;
  updatedAtIso: string;
  storageMode: GroupStorageMode;
  anchorEmailKey: string;
  includedEmailKeys: string[];
  workingGroupId?: string;
  workingGroupName?: string;
  filters: GroupWorksetFilterSnapshot;
  attachments: GroupPreparedAttachmentDescriptor[];
  mainLocation: GroupStorageLocationPointer;
  remotePromotionLocation?: GroupStorageLocationPointer | null;
  promotion: GroupWorksetPromotionStatus;
};

export type GroupPromotionPolicy = {
  mainPersistence: GroupStorageLocationPointer;
  remotePromotionLocation?: GroupStorageLocationPointer | null;
  allowRemotePromotion: boolean;
  promoteManifestOnPrimarySave: boolean;
  promoteAttachmentMetadataOnPrimarySave: boolean;
  promoteAttachmentBinaryOnPrimarySave: boolean;
  requireExplicitRemotePromotion: boolean;
  requireFreshPayloadBeforeOverwrite: boolean;
  saveSessionBeforeContextChange: boolean;
  saveSessionBeforeExit: boolean;
};

export type GroupAttachmentStoragePolicy = {
  thresholdBytes: number;
  ignoreInlineAttachments: boolean;
  mainBinaryStrategy: GroupAttachmentBinaryStrategy;
  remoteBinaryStrategy: GroupAttachmentBinaryStrategy;
  promptWhenAboveThreshold: boolean;
  neverAutoPromoteBinary: boolean;
};
