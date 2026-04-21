import type { IntermediateCase } from "./intermediateCaseTypes";
import { normalizeIntermediateCase } from "./intermediateCaseNormalization";

export function serializeIntermediateCase(caseValue: IntermediateCase): string {
  const normalized = normalizeIntermediateCase(caseValue);
  if (!normalized) {
    throw new Error("Intermediate case invalido: faltam caseId ou anchorEmailKey.");
  }
  return JSON.stringify(normalized, null, 2);
}

export function parseIntermediateCase(raw: string): IntermediateCase | null {
  try {
    const parsed = JSON.parse(String(raw || ""));
    return normalizeIntermediateCase(parsed);
  } catch {
    return null;
  }
}
