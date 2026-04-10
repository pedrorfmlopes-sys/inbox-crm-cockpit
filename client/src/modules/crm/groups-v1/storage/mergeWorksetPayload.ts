import { hasMeaningfulGroupWorksetPayload } from "./guards";
import { buildGroupWorksetManifest, normalizeGroupWorksetManifest } from "./worksetManifest";
import type { GroupWorksetManifest } from "./types";

export function mergeGroupWorksetPayload(
  current: GroupWorksetManifest | null | undefined,
  incoming: GroupWorksetManifest | null | undefined
): GroupWorksetManifest | null {
  const next = normalizeGroupWorksetManifest(incoming);
  if (!next) return normalizeGroupWorksetManifest(current);

  const previous = normalizeGroupWorksetManifest(current);
  if (!previous) return next;
  if (!hasMeaningfulGroupWorksetPayload(next) && hasMeaningfulGroupWorksetPayload(previous)) {
    return previous;
  }

  return buildGroupWorksetManifest({
    ...previous,
    ...next,
    createdAtIso: previous.createdAtIso,
    updatedAtIso: next.updatedAtIso,
    mainLocation: next.mainLocation || previous.mainLocation,
    remotePromotionLocation: next.remotePromotionLocation === undefined
      ? previous.remotePromotionLocation
      : next.remotePromotionLocation,
    promotion: {
      ...previous.promotion,
      ...next.promotion,
      promotedScopes: Array.isArray(next.promotion?.promotedScopes)
        ? next.promotion.promotedScopes
        : previous.promotion.promotedScopes,
      blockedScopes: Array.isArray(next.promotion?.blockedScopes)
        ? next.promotion.blockedScopes
        : previous.promotion.blockedScopes,
    },
  });
}
