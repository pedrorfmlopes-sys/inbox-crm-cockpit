export type GroupMembershipKind = "principal" | "referencia";
export type GroupTaskStatus = "por_fazer" | "em_curso" | "concluida" | "bloqueada" | "adiada";
export type GroupTaskPriority = "baixa" | "media" | "alta";
export type GroupChangeScope = "single_email";

export interface EmailGroupSelectionState {
  principalGroupId: string;
  referenceGroupIds: string[];
}

export interface GroupChangeContract {
  scope: GroupChangeScope;
  requiresExplicitChange: true;
  requiresWarning: true;
  allowsPreviousGroupAsReference: true;
  previousGroupReferenceScope: GroupChangeScope;
}

export interface GroupChangeRequest {
  emailKey: string;
  previousPrincipalGroupId: string | null;
  nextPrincipalGroupId: string | null;
  keepPreviousGroupAsReference: boolean;
  scope: GroupChangeScope;
}

export interface GroupTaskCoreFields {
  title: string;
  status: GroupTaskStatus;
  priority: GroupTaskPriority;
  dueDate?: string;
  owner?: string;
  originContext?: string;
  notes?: string;
}

export interface GroupTaskEntry extends GroupTaskCoreFields {
  id: string;
  groupId?: string;
  createdAt?: string;
  updatedAt?: string;
}

export interface GroupsPersistenceContract {
  sessionProgressStore: "session_cache";
  remotePersistenceTarget: "backend_link_store";
  remoteWritePolicy: "before_context_leave";
  avoidPrematureRemoteWrites: true;
  saveSessionBeforeContextLeave: true;
}

export const GROUPS_PRIMARY_GROUP_MAX_PER_EMAIL = 1;

export const GROUP_CHANGE_CONTRACT: GroupChangeContract = {
  scope: "single_email",
  requiresExplicitChange: true,
  requiresWarning: true,
  allowsPreviousGroupAsReference: true,
  previousGroupReferenceScope: "single_email",
};

export const GROUP_TASK_STATUS_ORDER: GroupTaskStatus[] = [
  "por_fazer",
  "em_curso",
  "concluida",
  "bloqueada",
  "adiada",
];

export const DEFAULT_GROUP_TASK_DRAFT: GroupTaskCoreFields = {
  title: "",
  status: "por_fazer",
  priority: "media",
  dueDate: "",
  owner: "",
  originContext: "",
  notes: "",
};

export const GROUPS_PERSISTENCE_CONTRACT: GroupsPersistenceContract = {
  sessionProgressStore: "session_cache",
  remotePersistenceTarget: "backend_link_store",
  remoteWritePolicy: "before_context_leave",
  avoidPrematureRemoteWrites: true,
  saveSessionBeforeContextLeave: true,
};

function normalizeGroupId(value: string | null | undefined): string {
  return String(value || "").trim();
}

export function normalizeGroupIdList(values: Array<string | null | undefined>): string[] {
  return Array.from(
    new Set(values.map((value) => normalizeGroupId(value)).filter(Boolean))
  );
}

export function createEmailGroupSelectionState(
  input: Partial<EmailGroupSelectionState> = {}
): EmailGroupSelectionState {
  const principalGroupId = normalizeGroupId(input.principalGroupId);
  const referenceGroupIds = normalizeGroupIdList(input.referenceGroupIds || []).filter(
    (groupId) => groupId !== principalGroupId
  );

  return {
    principalGroupId,
    referenceGroupIds,
  };
}

export function setPrincipalGroupSelection(
  current: Partial<EmailGroupSelectionState>,
  nextPrincipalGroupId: string | null | undefined
): EmailGroupSelectionState {
  return createEmailGroupSelectionState({
    principalGroupId: nextPrincipalGroupId,
    referenceGroupIds: current.referenceGroupIds || [],
  });
}

export function toggleReferenceGroupSelection(
  current: Partial<EmailGroupSelectionState>,
  groupId: string | null | undefined
): EmailGroupSelectionState {
  const selection = createEmailGroupSelectionState(current);
  const normalizedGroupId = normalizeGroupId(groupId);

  if (!normalizedGroupId || normalizedGroupId === selection.principalGroupId) {
    return selection;
  }

  const nextReferenceGroupIds = selection.referenceGroupIds.includes(normalizedGroupId)
    ? selection.referenceGroupIds.filter((entry) => entry !== normalizedGroupId)
    : [...selection.referenceGroupIds, normalizedGroupId];

  return createEmailGroupSelectionState({
    principalGroupId: selection.principalGroupId,
    referenceGroupIds: nextReferenceGroupIds,
  });
}

export function addReferenceGroupSelection(
  current: Partial<EmailGroupSelectionState>,
  groupId: string | null | undefined
): EmailGroupSelectionState {
  const selection = createEmailGroupSelectionState(current);
  const normalizedGroupId = normalizeGroupId(groupId);

  if (!normalizedGroupId || normalizedGroupId === selection.principalGroupId) {
    return selection;
  }

  return createEmailGroupSelectionState({
    principalGroupId: selection.principalGroupId,
    referenceGroupIds: selection.referenceGroupIds.includes(normalizedGroupId)
      ? selection.referenceGroupIds
      : [...selection.referenceGroupIds, normalizedGroupId],
  });
}

export function buildGroupChangeRequest(input: {
  emailKey: string | null | undefined;
  previousPrincipalGroupId?: string | null | undefined;
  nextPrincipalGroupId?: string | null | undefined;
  keepPreviousGroupAsReference?: boolean;
}): GroupChangeRequest | null {
  const emailKey = normalizeGroupId(input.emailKey);
  const previousPrincipalGroupId = normalizeGroupId(input.previousPrincipalGroupId);
  const nextPrincipalGroupId = normalizeGroupId(input.nextPrincipalGroupId);

  if (!emailKey || previousPrincipalGroupId === nextPrincipalGroupId) {
    return null;
  }

  return {
    emailKey,
    previousPrincipalGroupId: previousPrincipalGroupId || null,
    nextPrincipalGroupId: nextPrincipalGroupId || null,
    keepPreviousGroupAsReference: Boolean(
      input.keepPreviousGroupAsReference && previousPrincipalGroupId && previousPrincipalGroupId !== nextPrincipalGroupId
    ),
    scope: GROUP_CHANGE_CONTRACT.scope,
  };
}

export function shouldPromoteGroupsSessionProgress(input: {
  hasDirtySession: boolean;
  leavingContext: boolean;
  explicitRemoteSave?: boolean;
}): boolean {
  if (!input.hasDirtySession) return false;
  return Boolean(input.explicitRemoteSave || input.leavingContext);
}
