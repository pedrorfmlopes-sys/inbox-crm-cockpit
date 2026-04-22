import { getSettings } from "@/settings";
import { normalizeGroupsTabSettings } from "../settings/groupsTabSettings";
import { resolveIntermediateCaseStorage, type ResolvedIntermediateCaseStorage } from "./resolveIntermediateCaseStorage";
import type { IntermediateCase } from "./intermediateCaseTypes";

export type ResolvedClassificationIntermediateCase = {
  caseValue: IntermediateCase | null;
  lookup: "case_id" | "anchor_email_key" | "none";
  storage: ResolvedIntermediateCaseStorage;
};

export async function resolveClassificationIntermediateCase(input: {
  caseId?: string | null;
  anchorEmailKey?: string | null;
}): Promise<ResolvedClassificationIntermediateCase> {
  const settings = await getSettings().catch(() => null);
  const groupsSettings = normalizeGroupsTabSettings(settings?.groupsTabSettings || null);
  const storage = resolveIntermediateCaseStorage(groupsSettings);
  const caseId = String(input.caseId || "").trim();
  const anchorEmailKey = String(input.anchorEmailKey || "").trim();

  if (caseId) {
    const byCaseId = await storage.repository.readCase(caseId).catch(() => null);
    if (byCaseId) {
      return { caseValue: byCaseId, lookup: "case_id", storage };
    }
  }

  if (anchorEmailKey) {
    const byAnchor = await storage.repository.findCaseByEmailKey(anchorEmailKey).catch(() => null);
    if (byAnchor) {
      return { caseValue: byAnchor, lookup: "anchor_email_key", storage };
    }
  }

  return {
    caseValue: null,
    lookup: "none",
    storage,
  };
}
