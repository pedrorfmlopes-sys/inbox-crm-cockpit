import { getGroupWorksetManifest as getGroupWorksetManifestApi } from "./worksetApi";
import { buildGroupWorksetKey, supportsPrimaryGroupWorksetPersistence } from "./guards";
import { normalizeGroupWorksetManifest } from "./worksetManifest";
import type { ResolvedGroupStorageRuntime } from "./resolveStorageMode";
import type { GroupWorksetManifest } from "./types";

export async function loadPrimaryGroupWorkset(input: {
  anchorEmailKey: string;
  runtime: ResolvedGroupStorageRuntime;
}): Promise<GroupWorksetManifest | null> {
  if (!supportsPrimaryGroupWorksetPersistence(input.runtime.mode)) {
    return null;
  }
  const worksetKey = buildGroupWorksetKey(input.anchorEmailKey);
  if (!worksetKey) return null;
  const manifest = await getGroupWorksetManifestApi(worksetKey, {
    mode: input.runtime.mode,
    basePath: input.runtime.primaryLocation.basePath,
    chosenFolderKind: input.runtime.primaryLocation.kind === "document_library" ? "document_library" : "filesystem",
    primaryTarget: input.runtime.mode === "hybrid"
      ? input.runtime.settings.hybrid.primaryTarget
      : undefined,
  });
  return normalizeGroupWorksetManifest(manifest);
}
