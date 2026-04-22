import { saveGroupWorksetManifest as saveGroupWorksetManifestApi } from "./worksetApi";
import { hasMeaningfulGroupWorksetPayload, supportsPrimaryGroupWorksetPersistence } from "./guards";
import { mergeGroupWorksetPayload } from "./mergeWorksetPayload";
import type { ResolvedGroupStorageRuntime } from "./resolveStorageMode";
import type { GroupWorksetManifest } from "./types";

export async function savePrimaryGroupWorkset(input: {
  runtime: ResolvedGroupStorageRuntime;
  manifest: GroupWorksetManifest | null;
  current?: GroupWorksetManifest | null;
  keepalive?: boolean;
}): Promise<GroupWorksetManifest | null> {
  if (!input.manifest) return input.current || null;
  if (!supportsPrimaryGroupWorksetPersistence(input.runtime)) {
    return input.current || null;
  }
  const merged = mergeGroupWorksetPayload(input.current, input.manifest);
  if (!merged || !hasMeaningfulGroupWorksetPayload(merged)) {
    return input.current || null;
  }
  return await saveGroupWorksetManifestApi({
    manifest: merged,
    keepalive: input.keepalive,
  });
}
