import type { GroupTicketEntry, LinkGroupEntry, RelatedEmailEntry } from "@/api";
import type { GroupLabelCatalogEntry, CockpitSettingsV1 } from "@/settings";
import { clientLog } from "@/logger";
import {
  beginOutlookCategoryOperation,
  completeOutlookCategoryOperation,
  enqueueOutlookCategorySyncRequest,
  getManagedOutlookCategorySnapshot,
  requestCockpitHostAction,
  setOutlookCategoryOperationPhase,
  waitForOutlookCategorySyncResult,
  type OutlookCategorySyncTarget,
} from "@/office";
import {
  buildOutlookCategoryPlan,
  buildOutlookCategorySourceFromRelatedContext,
  getOutlookCategoryPlanSignature,
  getOutlookCategorySourceSignature,
} from "@/outlookCategories";

import {
  buildRemoteApplyFallbackCurrentCategoryEmail,
  type ApplyCurrentContext,
  type ResolvedStudioApplySelection,
} from "./applyResolution";
import { getApplyOperationErrorMessage } from "./applyOperationFinalization";

export type RefreshedClassificationContext = {
  email: RelatedEmailEntry | null;
  emails: RelatedEmailEntry[];
  groups: LinkGroupEntry[];
  tickets: GroupTicketEntry[];
};

export type BeginApplyOutlookCategoryOperationResult =
  | { ok: true; activeCategoryOperationId: string }
  | { ok: false; error: string };

export type PostApplyOutlookCategorySyncResult = {
  categoryOperationClosed: boolean;
};

export function beginApplyOutlookCategoryOperation(args: {
  currentTargetIdentity?: OutlookCategorySyncTarget | null;
}): BeginApplyOutlookCategoryOperationResult {
  if (!args.currentTargetIdentity) {
    return { ok: true, activeCategoryOperationId: "" };
  }

  const openedOperation = beginOutlookCategoryOperation({
    owner: "classification",
    target: args.currentTargetIdentity,
  });
  if (!openedOperation.ok) {
    return {
      ok: false,
      error: openedOperation.reason === "locked"
        ? "Ja existe outra classificacao em curso para este email. Aguarda um momento."
        : "Nao foi possivel identificar o email atual para confirmar a classificacao.",
    };
  }

  setOutlookCategoryOperationPhase(openedOperation.operation.operationId, "saving");
  return {
    ok: true,
    activeCategoryOperationId: openedOperation.operation.operationId,
  };
}

export async function executePostApplyOutlookCategorySync(args: {
  activeCategoryOperationId?: string;
  currentTargetIdentity?: OutlookCategorySyncTarget | null;
  includesCurrentTarget: boolean;
  effectiveTargetEmails: RelatedEmailEntry[];
  selectedEmail: RelatedEmailEntry | null;
  currentContext: ApplyCurrentContext;
  resolvedApplySelection: ResolvedStudioApplySelection;
  refreshedContext: RefreshedClassificationContext | null;
  latestSettings: Pick<CockpitSettingsV1, "groups"> | null | undefined;
  labelCatalog: GroupLabelCatalogEntry[];
  principalGroup: LinkGroupEntry | null;
  referenceGroups: LinkGroupEntry[];
  finalTicket: GroupTicketEntry | null;
  currentOutlookTicket: GroupTicketEntry | null;
  setStatus: (status: string) => void;
  logSync?: (phase: string, data: unknown) => void;
}): Promise<PostApplyOutlookCategorySyncResult> {
  const {
    activeCategoryOperationId,
    currentTargetIdentity,
    includesCurrentTarget,
    effectiveTargetEmails,
    selectedEmail,
    currentContext,
    resolvedApplySelection,
    refreshedContext,
    latestSettings,
    labelCatalog,
    principalGroup,
    referenceGroups,
    finalTicket,
    currentOutlookTicket,
    setStatus,
    logSync,
  } = args;

  if (!activeCategoryOperationId || !currentTargetIdentity || !includesCurrentTarget) {
    return { categoryOperationClosed: false };
  }

  let categoryOperationClosed = false;
  try {
    setOutlookCategoryOperationPhase(activeCategoryOperationId, "rehydrating");

    const currentTargetEmail = effectiveTargetEmails.find((email) => (
      String(email?.itemId || "").trim()
        ? String(email?.itemId || "").trim() === String(currentContext.itemId || "").trim()
        : String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
            === String(currentContext.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
    )) || (
      selectedEmail && (
        String(selectedEmail?.itemId || "").trim()
          ? String(selectedEmail?.itemId || "").trim() === String(currentContext.itemId || "").trim()
          : String(selectedEmail?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
              === String(currentContext.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
      )
        ? selectedEmail
        : null
    );

    const fallbackCurrentCategoryEmail = buildRemoteApplyFallbackCurrentCategoryEmail({
      currentTargetEmail,
      currentContext,
      resolvedApplySelection,
    });

    const refreshedCategoryEmailCandidates = [
      ...(refreshedContext?.email ? [refreshedContext.email] : []),
      ...(Array.isArray(refreshedContext?.emails) ? refreshedContext.emails : []),
      ...(fallbackCurrentCategoryEmail ? [fallbackCurrentCategoryEmail] : []),
    ].filter(Boolean) as RelatedEmailEntry[];

    const refreshedCategoryEmail = refreshedCategoryEmailCandidates.find((email) => (
      String(email?.itemId || "").trim()
        ? String(email?.itemId || "").trim() === String(currentContext.itemId || "").trim()
        : String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
            === String(currentContext.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
    )) || fallbackCurrentCategoryEmail;

    if (!refreshedCategoryEmail) {
      completeOutlookCategoryOperation(activeCategoryOperationId, {
        result: "failed",
        detail: "missing-refreshed-email",
      });
      categoryOperationClosed = true;
      throw new Error("A classificacao foi guardada, mas nao foi possivel rehidratar o email final para projetar as categorias.");
    }

    setOutlookCategoryOperationPhase(activeCategoryOperationId, "planning");
    const refreshedSnapshot = await getManagedOutlookCategorySnapshot(labelCatalog).catch(() => null);
    const refreshedCategorySource = buildOutlookCategorySourceFromRelatedContext({
      email: refreshedCategoryEmail,
      groups: Array.isArray(refreshedContext?.groups)
        ? refreshedContext.groups
        : [principalGroup, ...referenceGroups].filter(Boolean) as LinkGroupEntry[],
      tickets: Array.isArray(refreshedContext?.tickets)
        ? refreshedContext.tickets
        : [finalTicket, currentOutlookTicket].filter(Boolean) as GroupTicketEntry[],
      settings: latestSettings,
      currentOutlookLabelNames: refreshedSnapshot?.labelNames || [],
    });

    const categoryRequestId = `classification-final:${Date.now()}:${Math.random().toString(36).slice(2)}`;
    const categoryRequestedAtIso = new Date().toISOString();
    const categoryPlan = buildOutlookCategoryPlan(refreshedCategorySource);

    logSync?.("final-request", {
      requestId: categoryRequestId,
      operationId: activeCategoryOperationId || undefined,
      target: currentTargetIdentity,
      sourceSignature: getOutlookCategorySourceSignature(refreshedCategorySource),
      planSignature: getOutlookCategoryPlanSignature(categoryPlan),
      desiredCategories: categoryPlan.desiredCategories,
    });

    enqueueOutlookCategorySyncRequest({
      requestId: categoryRequestId,
      operationId: activeCategoryOperationId || undefined,
      createdAtIso: categoryRequestedAtIso,
      reason: "classification-final",
      mode: "source",
      target: currentTargetIdentity,
      source: refreshedCategorySource,
    });

    setOutlookCategoryOperationPhase(activeCategoryOperationId, "writingOutlook", {
      requestId: categoryRequestId,
    });

    setStatus("A projetar categorias no Outlook...");
    const writerSubmitted = await requestCockpitHostAction({
      type: "sync-managed-categories",
      payload: refreshedCategorySource,
      requestId: categoryRequestId,
      operationId: activeCategoryOperationId || undefined,
      requestedAtIso: categoryRequestedAtIso,
      reason: "classification-final",
      target: currentTargetIdentity,
    }).catch(() => false);

    if (!writerSubmitted) {
      completeOutlookCategoryOperation(activeCategoryOperationId, {
        result: "failed",
        requestId: categoryRequestId,
        detail: "writer-submit-failed",
      });
      categoryOperationClosed = true;
      throw new Error("A classificacao foi guardada, mas nao foi possivel submeter a projecao Outlook.");
    }

    setOutlookCategoryOperationPhase(activeCategoryOperationId, "verifying", {
      requestId: categoryRequestId,
    });

    const writerResult = await waitForOutlookCategorySyncResult(categoryRequestId, {
      timeoutMs: 20_000,
    });

    if (!writerResult) {
      clientLog.warn("[TEMP][outlook-category-apply] writer-timeout", {
        itemId: currentTargetIdentity.itemId,
        requestId: categoryRequestId,
        categoriesRequested: categoryPlan.desiredCategories,
        reason: "waitForOutlookCategorySyncResult timeout",
      });
      completeOutlookCategoryOperation(activeCategoryOperationId, {
        result: "timeout",
        requestId: categoryRequestId,
        detail: "writer-timeout",
      });
      categoryOperationClosed = true;
      throw new Error("A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias a tempo.");
    }

    completeOutlookCategoryOperation(activeCategoryOperationId, {
      result: writerResult.result,
      requestId: categoryRequestId,
      detail: writerResult.detail,
    });
    categoryOperationClosed = true;

    if (writerResult.result !== "success" && writerResult.result !== "duplicate") {
      const writerDetail = String(writerResult.detail || "").trim();
      clientLog.warn("[TEMP][outlook-category-apply] writer-degraded-result", {
        itemId: currentTargetIdentity.itemId,
        requestId: categoryRequestId,
        categoriesRequested: categoryPlan.desiredCategories,
        writerResult: writerResult.result,
        detail: writerDetail || undefined,
      });
      throw new Error(
        writerDetail
          ? `A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias (${writerDetail}).`
          : "A classificacao foi guardada, mas o Outlook nao confirmou a aplicacao das categorias."
      );
    }

    return { categoryOperationClosed };
  } catch (error) {
    if (activeCategoryOperationId && !categoryOperationClosed) {
      completeOutlookCategoryOperation(activeCategoryOperationId, {
        result: "failed",
        detail: getApplyOperationErrorMessage(error) || undefined,
      });
      categoryOperationClosed = true;
    }
    throw error;
  }
}
