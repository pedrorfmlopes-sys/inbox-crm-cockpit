import type { GroupTicketEntry, LinkGroupEntry, RelatedEmailEntry } from "@/api";
import type { CockpitSettingsV1 } from "@/settings";

export const ODOO_LINKED_CATEGORY = "Odoo Linked";
export const CRM_FOLLOW_UP_CATEGORY = "CRM: Follow-up";
export const GROUP_CATEGORY_PREFIX = "Grupo: ";
export const REFERENCE_CATEGORY_PREFIX = "Ref: ";
export const TICKET_CATEGORY_PREFIX = "TK: ";
export const LEGACY_STATUS_CATEGORY_PREFIX = "Estado: ";
export const GROUP_STATUS_CATEGORY_PREFIX = "Gr: ";
export const TICKET_STATUS_CATEGORY_PREFIX = "E-Tk: ";
export const LABEL_STATUS_CATEGORY_PREFIX = "E-Et: ";
export const LEGACY_TICKET_CATEGORY_PREFIX = "Ticket: ";
export const LEGACY_LABEL_CATEGORY_PREFIX = "Etiqueta: ";

const MANAGED_CATEGORY_PREFIXES = [
  GROUP_CATEGORY_PREFIX,
  REFERENCE_CATEGORY_PREFIX,
  TICKET_CATEGORY_PREFIX,
  LEGACY_STATUS_CATEGORY_PREFIX,
  GROUP_STATUS_CATEGORY_PREFIX,
  TICKET_STATUS_CATEGORY_PREFIX,
  LABEL_STATUS_CATEGORY_PREFIX,
  LEGACY_TICKET_CATEGORY_PREFIX,
  LEGACY_LABEL_CATEGORY_PREFIX,
];

const RESERVED_SPECIAL_CATEGORY_NAMES = new Set([
  ODOO_LINKED_CATEGORY.toLowerCase(),
  CRM_FOLLOW_UP_CATEGORY.toLowerCase(),
]);

export type OutlookCategorySource = {
  principalGroupNames: string[];
  referenceGroupNames: string[];
  ticketCodes: string[];
  labelNames: string[];
  managedLabelNames: string[];
  groupStatuses: string[];
  ticketStatuses: string[];
  labelStatuses: string[];
  specialCategories: string[];
  managedSpecialCategories: string[];
};

export type LegacyManagedOutlookCategoryInput = {
  principalGroupNames?: string[];
  referenceGroupNames?: string[];
  groupNames?: string[];
  ticketCodes?: string[];
  statuses?: string[];
  groupStatuses?: string[];
  ticketStatuses?: string[];
  labelStatuses?: string[];
  labelNames?: string[];
  managedLabelNames?: string[];
  specialCategories?: string[];
  managedSpecialCategories?: string[];
};

export type OutlookCategoryPlan = {
  source: OutlookCategorySource;
  desiredCategories: string[];
  managedLabelNames: string[];
  managedSpecialCategories: string[];
  manageClassificationFamilies: boolean;
};

export function normalizeUniqueCategoryValues(values?: readonly string[] | null): string[] {
  return Array.from(new Set((values || []).map((value) => String(value || "").trim()).filter(Boolean)));
}

export function normalizeGroupStatusCategoryLabel(value: string | undefined): string {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "concluido") return "Concluido";
  if (normalized === "em_progresso") return "Em progresso";
  if (normalized === "em_analise") return "Em analise";
  return String(value || "").trim();
}

export function normalizeTicketStatusCategoryLabel(value: string | undefined): string {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "open" || normalized === "aberto") return "Aberto";
  if (normalized === "closed" || normalized === "fechado") return "Fechado";
  return normalizeGroupStatusCategoryLabel(value);
}

export function normalizeOutlookCategorySource(source?: Partial<OutlookCategorySource> | null): OutlookCategorySource {
  return {
    principalGroupNames: normalizeUniqueCategoryValues(source?.principalGroupNames),
    referenceGroupNames: normalizeUniqueCategoryValues(source?.referenceGroupNames),
    ticketCodes: normalizeUniqueCategoryValues(source?.ticketCodes),
    labelNames: normalizeUniqueCategoryValues(source?.labelNames),
    managedLabelNames: normalizeUniqueCategoryValues(source?.managedLabelNames ?? source?.labelNames),
    groupStatuses: normalizeUniqueCategoryValues(source?.groupStatuses),
    ticketStatuses: normalizeUniqueCategoryValues(source?.ticketStatuses),
    labelStatuses: normalizeUniqueCategoryValues(source?.labelStatuses),
    specialCategories: normalizeUniqueCategoryValues(source?.specialCategories),
    managedSpecialCategories: normalizeUniqueCategoryValues(source?.managedSpecialCategories ?? source?.specialCategories),
  };
}

export function mergeOutlookCategorySources(
  ...sources: Array<Partial<OutlookCategorySource> | null | undefined>
): OutlookCategorySource {
  return normalizeOutlookCategorySource({
    principalGroupNames: sources.flatMap((source) => source?.principalGroupNames || []),
    referenceGroupNames: sources.flatMap((source) => source?.referenceGroupNames || []),
    ticketCodes: sources.flatMap((source) => source?.ticketCodes || []),
    labelNames: sources.flatMap((source) => source?.labelNames || []),
    managedLabelNames: sources.flatMap((source) => source?.managedLabelNames || []),
    groupStatuses: sources.flatMap((source) => source?.groupStatuses || []),
    ticketStatuses: sources.flatMap((source) => source?.ticketStatuses || []),
    labelStatuses: sources.flatMap((source) => source?.labelStatuses || []),
    specialCategories: sources.flatMap((source) => source?.specialCategories || []),
    managedSpecialCategories: sources.flatMap((source) => source?.managedSpecialCategories || []),
  });
}

function sortNormalizedCategoryValues(values: readonly string[]): string[] {
  return [...values].sort((left, right) => {
    const normalizedLeft = String(left || "").trim().toLowerCase();
    const normalizedRight = String(right || "").trim().toLowerCase();
    if (normalizedLeft !== normalizedRight) return normalizedLeft.localeCompare(normalizedRight, "pt");
    return String(left || "").trim().localeCompare(String(right || "").trim(), "pt");
  });
}

export function getOutlookCategorySourceSignature(
  source?: Partial<OutlookCategorySource> | null
): string {
  const normalized = normalizeOutlookCategorySource(source);
  return JSON.stringify({
    principalGroupNames: sortNormalizedCategoryValues(normalized.principalGroupNames),
    referenceGroupNames: sortNormalizedCategoryValues(normalized.referenceGroupNames),
    ticketCodes: sortNormalizedCategoryValues(normalized.ticketCodes),
    labelNames: sortNormalizedCategoryValues(normalized.labelNames),
    managedLabelNames: sortNormalizedCategoryValues(normalized.managedLabelNames),
    groupStatuses: sortNormalizedCategoryValues(normalized.groupStatuses),
    ticketStatuses: sortNormalizedCategoryValues(normalized.ticketStatuses),
    labelStatuses: sortNormalizedCategoryValues(normalized.labelStatuses),
    specialCategories: sortNormalizedCategoryValues(normalized.specialCategories),
    managedSpecialCategories: sortNormalizedCategoryValues(normalized.managedSpecialCategories),
  });
}

export function areOutlookCategorySourcesEqual(
  left?: Partial<OutlookCategorySource> | null,
  right?: Partial<OutlookCategorySource> | null
): boolean {
  return getOutlookCategorySourceSignature(left) === getOutlookCategorySourceSignature(right);
}

export function isManagedCategoryFamilyName(name: string): boolean {
  const label = String(name || "").trim();
  return MANAGED_CATEGORY_PREFIXES.some((prefix) => label.startsWith(prefix));
}

export function isReservedOutlookCategoryName(name: string): boolean {
  const normalized = String(name || "").trim().toLowerCase();
  return isManagedCategoryFamilyName(name) || RESERVED_SPECIAL_CATEGORY_NAMES.has(normalized);
}

export function buildOutlookCategoryPlan(
  source?: Partial<OutlookCategorySource> | null,
  options?: { manageClassificationFamilies?: boolean }
): OutlookCategoryPlan {
  const normalizedSource = normalizeOutlookCategorySource(source);
  const desiredCategories = [
    ...normalizedSource.principalGroupNames.map((name) => `${GROUP_CATEGORY_PREFIX}${name}`),
    ...normalizedSource.referenceGroupNames.map((name) => `${REFERENCE_CATEGORY_PREFIX}${name}`),
    ...normalizedSource.ticketCodes.map((code) => `${TICKET_CATEGORY_PREFIX}${code}`),
    ...normalizedSource.groupStatuses
      .map((status) => normalizeGroupStatusCategoryLabel(status))
      .filter(Boolean)
      .map((status) => `${GROUP_STATUS_CATEGORY_PREFIX}${status}`),
    ...normalizedSource.ticketStatuses
      .map((status) => normalizeTicketStatusCategoryLabel(status))
      .filter(Boolean)
      .map((status) => `${TICKET_STATUS_CATEGORY_PREFIX}${status}`),
    ...normalizedSource.labelStatuses
      .map((status) => normalizeGroupStatusCategoryLabel(status))
      .filter(Boolean)
      .map((status) => `${LABEL_STATUS_CATEGORY_PREFIX}${status}`),
    ...normalizedSource.labelNames,
    ...normalizedSource.specialCategories,
  ];
  return {
    source: normalizedSource,
    desiredCategories: normalizeUniqueCategoryValues(desiredCategories),
    managedLabelNames: normalizedSource.managedLabelNames,
    managedSpecialCategories: normalizedSource.managedSpecialCategories,
    manageClassificationFamilies: options?.manageClassificationFamilies !== false,
  };
}

export function getOutlookCategoryPlanSignature(plan: OutlookCategoryPlan): string {
  return JSON.stringify({
    desiredCategories: sortNormalizedCategoryValues(normalizeUniqueCategoryValues(plan.desiredCategories)),
    managedLabelNames: sortNormalizedCategoryValues(normalizeUniqueCategoryValues(plan.managedLabelNames)),
    managedSpecialCategories: sortNormalizedCategoryValues(normalizeUniqueCategoryValues(plan.managedSpecialCategories)),
    manageClassificationFamilies: plan.manageClassificationFamilies !== false,
  });
}

export function buildOutlookCategorySourceFromLegacyInput(input?: LegacyManagedOutlookCategoryInput | null): OutlookCategorySource {
  return normalizeOutlookCategorySource({
    principalGroupNames: input?.principalGroupNames ?? input?.groupNames,
    referenceGroupNames: input?.referenceGroupNames,
    ticketCodes: input?.ticketCodes,
    labelNames: input?.labelNames,
    managedLabelNames: input?.managedLabelNames ?? input?.labelNames,
    groupStatuses: input?.groupStatuses ?? input?.statuses,
    ticketStatuses: input?.ticketStatuses,
    labelStatuses: input?.labelStatuses,
    specialCategories: input?.specialCategories,
    managedSpecialCategories: input?.managedSpecialCategories ?? input?.specialCategories,
  });
}

function resolveRelatedCustomGroups(email: RelatedEmailEntry | null, groups: LinkGroupEntry[]): {
  principalGroups: LinkGroupEntry[];
  referenceGroups: LinkGroupEntry[];
  customGroups: LinkGroupEntry[];
} {
  const customGroupsById = new Map(
    (Array.isArray(groups) ? groups : [])
      .filter((group) => String(group?.kind || "").trim().toLowerCase() === "custom")
      .map((group) => [String(group?.id || "").trim(), group] as const)
      .filter(([id]) => Boolean(id))
  );
  const relatedCustomGroups = Array.isArray(email?.relatedGroups)
    ? email.relatedGroups
        .filter((group) => {
          const groupId = String(group?.id || "").trim();
          const resolved = groupId ? customGroupsById.get(groupId) : null;
          const kind = String(group?.kind || resolved?.kind || "").trim().toLowerCase();
          return kind === "custom";
        })
        .map((group) => {
          const groupId = String(group?.id || "").trim();
          const resolved = groupId ? customGroupsById.get(groupId) : null;
          return resolved || ({
            id: groupId,
            name: String(group?.name || "").trim(),
            status: "",
            labels: [],
            kind: "custom",
          } as LinkGroupEntry);
        })
    : [];
  const principalGroups = relatedCustomGroups.filter(
    (group) => String((group as any)?.relationKind || "").trim().toLowerCase() === "principal"
  );
  const referenceGroups = relatedCustomGroups.filter(
    (group) => String((group as any)?.relationKind || "").trim().toLowerCase() === "referencia"
  );
  const customGroups = Array.from(new Map(relatedCustomGroups.map((group) => [String(group.id || "").trim(), group])).values());
  return { principalGroups, referenceGroups, customGroups };
}

function includesLabel(labelNames: string[], label: string): boolean {
  const normalized = String(label || "").trim().toLowerCase();
  return Boolean(normalized) && labelNames.some((entry) => entry.toLowerCase() === normalized);
}

function buildEffectiveLabels(email: RelatedEmailEntry | null, principalGroups: LinkGroupEntry[], referenceGroups: LinkGroupEntry[]) {
  const inheritedLabels = normalizeUniqueCategoryValues(
    [...principalGroups, ...referenceGroups]
      .flatMap((group) => Array.isArray(group?.labels) ? group.labels : [])
      .map((label) => String(label || "").trim())
  );
  const removedInheritedLabels = normalizeUniqueCategoryValues(email?.removedInheritedLabels);
  const directLabels = normalizeUniqueCategoryValues(email?.labels);
  const effectiveLabels = normalizeUniqueCategoryValues([
    ...inheritedLabels.filter((label) => !includesLabel(removedInheritedLabels, label)),
    ...directLabels,
  ]);
  const explicitCategorizedLabelNames = Array.isArray(email?.classificationMeta?.categorizedLabelNames)
    ? normalizeUniqueCategoryValues(email?.classificationMeta?.categorizedLabelNames).filter((label) => includesLabel(effectiveLabels, label))
    : [];
  const categorizedLabelNames = explicitCategorizedLabelNames.length ? explicitCategorizedLabelNames : effectiveLabels;
  const labelStatuses = email?.labelStates && typeof email.labelStates === "object"
    ? normalizeUniqueCategoryValues(
        Object.entries(email.labelStates)
          .filter(([label, value]) => Boolean(String(label || "").trim()) && Boolean(String(value || "").trim()) && includesLabel(effectiveLabels, label))
          .map(([, value]) => String(value || "").trim())
      )
    : [];
  return {
    effectiveLabels,
    removedInheritedLabels,
    categorizedLabelNames,
    labelStatuses,
  };
}

export function buildOutlookCategorySourceFromRelatedContext(input: {
  email: RelatedEmailEntry | null;
  groups: LinkGroupEntry[];
  tickets: GroupTicketEntry[];
  settings: Pick<CockpitSettingsV1, "groupOutlookCategories"> | null | undefined;
  currentOutlookLabelNames?: string[];
  specialCategories?: string[];
  managedSpecialCategories?: string[];
}): OutlookCategorySource {
  const email = input.email;
  const categorySettings = input.settings?.groupOutlookCategories;
  const categoriesEnabled = categorySettings?.enabled === true;
  const includeGroups = categoriesEnabled && categorySettings?.includeGroups !== false;
  const includeTickets = categoriesEnabled && categorySettings?.includeTickets !== false;
  const includeStatuses = categoriesEnabled && categorySettings?.includeStatuses !== false;
  const includeLabels = categoriesEnabled && categorySettings?.includeLabels === true;
  const { principalGroups, referenceGroups, customGroups } = resolveRelatedCustomGroups(email, input.groups);
  const principalGroupNames = normalizeUniqueCategoryValues(
    principalGroups.map((group) => String(group?.name || "").trim())
  );
  const referenceGroupNames = normalizeUniqueCategoryValues(
    referenceGroups.map((group) => String(group?.name || "").trim())
  );
  const fallbackGroupNames = normalizeUniqueCategoryValues(
    customGroups.map((group) => String(group?.name || "").trim())
  );
  const effectivePrincipalGroupNames = principalGroupNames.length || referenceGroupNames.length
    ? principalGroupNames
    : fallbackGroupNames;
  const principalCategorize = email?.classificationMeta?.principalCategorize !== false;
  const referenceCategorize = email?.classificationMeta?.referenceCategorize !== false;
  const { effectiveLabels, removedInheritedLabels, categorizedLabelNames, labelStatuses } = buildEffectiveLabels(
    email,
    principalGroups,
    referenceGroups
  );
  const groupStatuses = normalizeUniqueCategoryValues([
    ...(email?.classificationMeta?.principalStatusEnabled && email?.classificationMeta?.principalStatusCategorize
      ? principalGroups.map((group) => String(group?.status || "").trim())
      : []),
    ...(email?.classificationMeta?.referenceStatusEnabled && email?.classificationMeta?.referenceStatusCategorize
      ? referenceGroups.map((group) => String(group?.status || "").trim())
      : []),
  ]);
  const ticketCodes = normalizeUniqueCategoryValues(
    (Array.isArray(input.tickets) ? input.tickets : []).map((ticket) => String(ticket?.code || "").trim())
  );
  const ticketStatuses = email?.classificationMeta?.ticketStatusEnabled && email?.classificationMeta?.ticketStatusCategorize
    ? normalizeUniqueCategoryValues(
        (Array.isArray(input.tickets) ? input.tickets : []).map((ticket) => String(ticket?.status || "").trim())
      )
    : [];
  return normalizeOutlookCategorySource({
    principalGroupNames: includeGroups && principalCategorize ? effectivePrincipalGroupNames : [],
    referenceGroupNames: includeGroups && referenceCategorize ? referenceGroupNames : [],
    ticketCodes: includeTickets ? ticketCodes : [],
    labelNames: includeLabels ? categorizedLabelNames : [],
    managedLabelNames: [
      ...effectiveLabels,
      ...removedInheritedLabels,
      ...normalizeUniqueCategoryValues(input.currentOutlookLabelNames),
    ],
    groupStatuses: includeStatuses ? groupStatuses : [],
    ticketStatuses: includeStatuses ? ticketStatuses : [],
    labelStatuses: includeStatuses ? labelStatuses : [],
    specialCategories: input.specialCategories,
    managedSpecialCategories: input.managedSpecialCategories,
  });
}
