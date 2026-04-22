import { completeOutlookCategoryOperation } from "@/office";

export type ApplyOperationResult = {
  ok: boolean;
  coreSuccess: boolean;
  error?: string;
};

export function finalizeSuccessfulApplyOperation(args: {
  effectiveTargetCount: number;
  currentSignature: string;
  setStatus: (status: string) => void;
  setLastAppliedSignature: (signature: string) => void;
}): ApplyOperationResult {
  args.setStatus(
    args.effectiveTargetCount > 1
      ? `Classificacao aplicada a ${args.effectiveTargetCount} emails.`
      : "Classificacao aplicada ao email selecionado."
  );
  args.setLastAppliedSignature(args.currentSignature);
  return { ok: true, coreSuccess: true };
}

export function closeApplyOutlookOperationSafely(args: {
  activeCategoryOperationId?: string;
  categoryOperationClosed: boolean;
  actionError: unknown;
}): boolean {
  if (!args.activeCategoryOperationId || args.categoryOperationClosed) {
    return args.categoryOperationClosed;
  }

  const detail = getApplyOperationErrorMessage(args.actionError);
  completeOutlookCategoryOperation(args.activeCategoryOperationId, {
    result: "failed",
    detail: detail || undefined,
  });
  return true;
}

export function finalizeFailedApplyOperation(args: {
  activeCategoryOperationId?: string;
  categoryOperationClosed: boolean;
  actionError: unknown;
  setStatus: (status: string) => void;
  coreSuccess: boolean;
}): ApplyOperationResult {
  closeApplyOutlookOperationSafely({
    activeCategoryOperationId: args.activeCategoryOperationId,
    categoryOperationClosed: args.categoryOperationClosed,
    actionError: args.actionError,
  });

  const errorMsg = getApplyOperationErrorMessage(args.actionError)
    || "Nao foi possivel aplicar a classificacao.";
  args.setStatus(errorMsg);

  if (args.coreSuccess) {
    return { ok: true, coreSuccess: true, error: `Guardado com avisos: ${errorMsg}` };
  }

  return { ok: false, coreSuccess: false, error: errorMsg };
}

export function getApplyOperationErrorMessage(actionError: unknown): string {
  if (actionError instanceof Error) {
    return String(actionError.message || "").trim();
  }

  if (typeof actionError === "string") {
    return actionError.trim();
  }

  if (actionError && typeof actionError === "object" && "message" in actionError) {
    return String((actionError as { message?: unknown }).message || "").trim();
  }

  return "";
}
