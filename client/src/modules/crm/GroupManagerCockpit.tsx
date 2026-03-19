import React, { useEffect, useMemo, useState } from "react";
import {
  addEmailToLinkGroup,
  createLinkGroup,
  deleteLinkGroup,
  getGroupEmails,
  listLinkGroups,
  removeEmailFromLinkGroup,
  searchKnownEmails,
  updateLinkGroup,
  type LinkGroupEntry,
  type RelatedEmailEntry,
} from "@/api";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { openGroupExplorer, openLinkedOutlookEmail } from "@/office";
import { saveSettings } from "@/settings";
import { HelpHint } from "@/ui/HelpHint";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type GroupManagerView = "groups" | "detail" | "library" | "settings" | "labels";
type GroupStatusFilter = "all" | "em_analise" | "em_progresso" | "concluido";
type GroupArchiveFilter = "active" | "archived" | "all";
type MembershipKind = "principal" | "referencia";

type GroupDraft = {
  name: string;
  description: string;
  status: "em_analise" | "em_progresso" | "concluido";
  labelsText: string;
  documentsEnabled: boolean;
  isArchived: boolean;
};

const STATUS_OPTIONS: Array<{ value: GroupDraft["status"]; label: string }> = [
  { value: "em_analise", label: "Em analise" },
  { value: "em_progresso", label: "Em progresso" },
  { value: "concluido", label: "Concluido" },
];

const MEMBERSHIP_OPTIONS: Array<{ value: MembershipKind; label: string }> = [
  { value: "principal", label: "Principal" },
  { value: "referencia", label: "Referencia" },
];

function normalizeText(value: string | undefined): string {
  return String(value || "").trim().toLowerCase();
}

function parseLabels(value: string | string[] | undefined): string[] {
  const raw = Array.isArray(value)
    ? value
    : String(value || "")
      .split(/[,\n;]/g)
      .map((entry) => entry.trim());
  const seen = new Set<string>();
  const labels: string[] = [];
  for (const item of raw) {
    const label = String(item || "").trim();
    if (!label) continue;
    const key = label.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    labels.push(label);
  }
  return labels.sort((a, b) => a.localeCompare(b, "pt-PT"));
}

function mergeLabelCatalog(...sources: Array<string[] | undefined>): string[] {
  const seen = new Map<string, string>();
  for (const source of sources) {
    for (const raw of source || []) {
      const label = String(raw || "").trim();
      if (!label) continue;
      const key = label.toLowerCase();
      if (!seen.has(key)) seen.set(key, label);
    }
  }
  return Array.from(seen.values()).sort((a, b) => a.localeCompare(b, "pt-PT"));
}

function canonicalizeLabelsWithCatalog(labels: string[], catalog: string[]): string[] {
  const map = new Map<string, string>();
  for (const label of catalog || []) {
    const normalized = String(label || "").trim();
    if (!normalized) continue;
    map.set(normalized.toLowerCase(), normalized);
  }
  return parseLabels(labels.map((label) => map.get(String(label || "").trim().toLowerCase()) || label));
}

function labelsToText(labels: string[] | undefined): string {
  return (labels || []).join(", ");
}

function statusLabel(value: string | undefined): string {
  if (value === "concluido") return "Concluido";
  if (value === "em_progresso") return "Em progresso";
  return "Em analise";
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function membershipKindLabel(value: string | undefined): string {
  return String(value || "").trim().toLowerCase() === "referencia" ? "Referencia" : "Principal";
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return (
    String(email?.emailKey || "").trim()
    || String(email?.id || "").trim()
    || String(email?.itemId || "").trim()
    || String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "")
    || [
      String(email?.conversationId || "").trim(),
      String(email?.subject || "").trim().toLowerCase(),
      String(email?.fromEmail || "").trim().toLowerCase(),
      String(email?.messageDateIso || email?.receivedAtIso || "").trim(),
    ].join("|")
  );
}

function createDraft(group: LinkGroupEntry | null): GroupDraft {
  return {
    name: String(group?.name || "").trim(),
    description: String(group?.description || "").trim(),
    status: group?.status === "em_progresso" || group?.status === "concluido" ? group.status : "em_analise",
    labelsText: labelsToText(group?.labels),
    documentsEnabled: group?.documentsEnabled !== false,
    isArchived: group?.isArchived === true,
  };
}

function isCurrentContextEmail(email: Partial<RelatedEmailEntry>, ctx: ReturnType<typeof useCockpit>["ctx"]): boolean {
  const currentItemId = String(ctx.itemId || "").trim();
  const emailItemId = String(email.itemId || "").trim();
  if (currentItemId && emailItemId && currentItemId === emailItemId) return true;

  const currentMessageId = String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  const emailMessageId = String(email.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
  if (currentMessageId && emailMessageId && currentMessageId === emailMessageId) return true;

  return Boolean(
    String(ctx.conversationId || "").trim()
    && String(email.conversationId || "").trim()
    && String(ctx.conversationId || "").trim() === String(email.conversationId || "").trim()
    && normalizeText(ctx.subject) === normalizeText(email.subject)
  );
}

function draftChanged(group: LinkGroupEntry | null, draft: GroupDraft): boolean {
  if (!group) return false;
  return (
    String(group.name || "") !== draft.name
    || String(group.description || "") !== draft.description
    || String(group.status || "em_analise") !== draft.status
    || labelsToText(group.labels) !== labelsToText(parseLabels(draft.labelsText))
    || (group.documentsEnabled !== false) !== draft.documentsEnabled
    || (group.isArchived === true) !== draft.isArchived
  );
}

const LabelPicker: React.FC<{
  value: string;
  onChange: (next: string) => void;
  catalog: string[];
  enabled: boolean;
  busy: boolean;
  placeholder?: string;
  helpText?: string;
  onCreateLabel: (label: string) => Promise<string | null>;
}> = ({ value, onChange, catalog, enabled, busy, placeholder, helpText, onCreateLabel }) => {
  const [query, setQuery] = useState("");
  const [open, setOpen] = useState(false);

  const selectedLabels = useMemo(() => parseLabels(value), [value]);

  const suggestions = useMemo(() => {
    const q = normalizeText(query);
    return (catalog || [])
      .filter((label) => !selectedLabels.some((entry) => normalizeText(entry) === normalizeText(label)))
      .filter((label) => !q || normalizeText(label).includes(q))
      .slice(0, 8);
  }, [catalog, query, selectedLabels]);

  const exactCatalogMatch = useMemo(
    () => (catalog || []).some((label) => normalizeText(label) === normalizeText(query)),
    [catalog, query]
  );

  const canCreate = enabled && Boolean(String(query || "").trim()) && !exactCatalogMatch && !suggestions.length;

  function commitLabel(label: string) {
    onChange(labelsToText(canonicalizeLabelsWithCatalog([...selectedLabels, label], catalog)));
    setQuery("");
    setOpen(false);
  }

  async function handleAction() {
    if (canCreate) {
      const created = await onCreateLabel(query);
      if (created) commitLabel(created);
      return;
    }
    setOpen(true);
  }

  return (
    <div style={S.labelPicker}>
      <div style={S.sectionTitleRow}>
        <div style={S.fieldLabel}>Etiquetas</div>
        {helpText ? <HelpHint text={helpText} title="Ajuda: Etiquetas" /> : null}
      </div>
      <div style={S.labelSearchRow}>
        <input
          style={S.input}
          value={query}
          onChange={(event) => {
            setQuery(event.target.value);
            setOpen(true);
          }}
          onFocus={() => setOpen(true)}
          onBlur={() => window.setTimeout(() => setOpen(false), 120)}
          placeholder={placeholder || "Pesquisar etiqueta"}
        />
        <button
          type="button"
          style={canCreate ? S.primaryBtn : S.iconGhostBtn}
          onClick={() => void handleAction()}
          disabled={busy || (!canCreate && !suggestions.length)}
          title={canCreate ? "Criar etiqueta e adicionar ao grupo" : "Pesquisar etiquetas existentes"}
        >
          {canCreate ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
        </button>
      </div>
      {open && suggestions.length ? (
        <div style={S.labelSuggestionList}>
          {suggestions.map((label) => (
            <button
              key={label}
              type="button"
              style={S.labelSuggestion}
              onMouseDown={(event) => {
                event.preventDefault();
                commitLabel(label);
              }}
            >
              {label}
            </button>
          ))}
        </div>
      ) : null}
      {selectedLabels.length ? (
        <div style={S.selectedLabelRow}>
          {selectedLabels.map((label) => (
            <button
              key={label}
              type="button"
              style={S.selectedLabelChip}
              onClick={() =>
                onChange(labelsToText(selectedLabels.filter((entry) => normalizeText(entry) !== normalizeText(label))))
              }
              title="Remover etiqueta"
            >
              <span>{label}</span>
              <span style={S.selectedLabelRemove}>×</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
};

export const GroupManagerCockpit: React.FC = () => {
  const { ctx, bodyText, bodyHtml, attachments, setMsg, setActiveGroupForCurrentEmail, settings, openSettingsSection } = useCockpit();
  const [view, setView] = useState<GroupManagerView>("groups");
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [groupsLoading, setGroupsLoading] = useState(false);
  const [groupsError, setGroupsError] = useState("");
  const [groupQuery, setGroupQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState<GroupStatusFilter>("all");
  const [archiveFilter, setArchiveFilter] = useState<GroupArchiveFilter>("active");
  const [activeLabelFilters, setActiveLabelFilters] = useState<string[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState("");
  const [draft, setDraft] = useState<GroupDraft>(createDraft(null));
  const [busy, setBusy] = useState(false);
  const [reloadToken, setReloadToken] = useState(0);

  const [groupEmails, setGroupEmails] = useState<RelatedEmailEntry[]>([]);
  const [groupEmailsLoading, setGroupEmailsLoading] = useState(false);
  const [groupEmailQuery, setGroupEmailQuery] = useState("");
  const [selectedGroupEmailKeys, setSelectedGroupEmailKeys] = useState<string[]>([]);
  const [linkKind, setLinkKind] = useState<MembershipKind>("principal");

  const [libraryQuery, setLibraryQuery] = useState("");
  const [libraryEmails, setLibraryEmails] = useState<RelatedEmailEntry[]>([]);
  const [libraryLoading, setLibraryLoading] = useState(false);
  const [selectedLibraryKeys, setSelectedLibraryKeys] = useState<string[]>([]);
  const [newCatalogLabel, setNewCatalogLabel] = useState("");
  const [selectedManagedLabel, setSelectedManagedLabel] = useState("");
  const [renameLabelValue, setRenameLabelValue] = useState("");

  const selectedGroup = useMemo(
    () => groups.find((group) => group.id === selectedGroupId) || null,
    [groups, selectedGroupId]
  );

  const currentEmailPayload = useMemo(
    () => ({
      itemId: String(ctx.itemId || "").trim(),
      internetMessageId: String(ctx.internetMessageId || "").trim(),
      conversationId: String(ctx.conversationId || "").trim(),
      subject: String(ctx.subject || "").trim(),
      fromEmail: String(ctx.fromEmail || "").trim(),
      fromName: String(ctx.fromName || "").trim(),
      receivedAtIso: String(ctx.receivedDateTimeIso || "").trim(),
      messageDateIso: String(ctx.receivedDateTimeIso || "").trim(),
      bodyText: String(bodyText || "").trim(),
      bodyHtml: String(bodyHtml || "").trim(),
      attachments: (attachments || []).map((attachment) => ({
        id: attachment.id,
        name: attachment.name,
        contentType: attachment.contentType,
        size: attachment.size,
        isInline: attachment.isInline,
        contentId: attachment.contentId,
        content: attachment.content,
      })),
    }),
    [attachments, bodyHtml, bodyText, ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject]
  );

  const currentEmailKey = useMemo(() => makeEmailKey(currentEmailPayload), [currentEmailPayload]);
  const favoriteGroupIds = settings?.groupFavoriteIds || [];
  const favoriteGroupSet = useMemo(() => new Set(favoriteGroupIds), [favoriteGroupIds]);

  useEffect(() => {
    let cancelled = false;
    setGroupsLoading(true);
    setGroupsError("");
    listLinkGroups("/")
      .then((rows) => {
        if (cancelled) return;
        setGroups(Array.isArray(rows) ? rows : []);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setGroups([]);
        setGroupsError(error?.message || "Nao foi possivel carregar grupos.");
      })
      .finally(() => {
        if (!cancelled) setGroupsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken]);

  useEffect(() => {
    setDraft(createDraft(selectedGroup));
  }, [selectedGroup]);

  useEffect(() => {
    if (!selectedGroupId) {
      setGroupEmails([]);
      setSelectedGroupEmailKeys([]);
      return;
    }
    let cancelled = false;
    setGroupEmailsLoading(true);
    getGroupEmails(selectedGroupId)
      .then((rows) => {
        if (cancelled) return;
        setGroupEmails(Array.isArray(rows) ? rows : []);
        setSelectedGroupEmailKeys([]);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setGroupEmails([]);
        setMsg(error?.message || "Nao foi possivel carregar os emails do grupo.");
      })
      .finally(() => {
        if (!cancelled) setGroupEmailsLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [reloadToken, selectedGroupId, setMsg]);

  useEffect(() => {
    let cancelled = false;
    setLibraryLoading(true);
    searchKnownEmails(libraryQuery, { excludeGroupId: selectedGroupId || undefined, limit: 120 })
      .then((rows) => {
        if (cancelled) return;
        setLibraryEmails(Array.isArray(rows) ? rows : []);
        setSelectedLibraryKeys([]);
      })
      .catch((error: any) => {
        if (cancelled) return;
        setLibraryEmails([]);
        setMsg(error?.message || "Nao foi possivel pesquisar emails registados.");
      })
      .finally(() => {
        if (!cancelled) setLibraryLoading(false);
      });
    return () => {
      cancelled = true;
    };
  }, [libraryQuery, reloadToken, selectedGroupId, setMsg]);

  const allLabels = useMemo(() => {
    return mergeLabelCatalog(
      settings?.groupLabelCatalog || [],
      groups.flatMap((group) => group.labels || [])
    );
  }, [groups, settings?.groupLabelCatalog]);

  const labelsManagerEnabled = settings?.groupLabelsManagerEnabled !== false;

  const labelUsage = useMemo(
    () =>
      allLabels.map((label) => ({
        label,
        count: groups.filter((group) =>
          (group.labels || []).some((entry) => normalizeText(entry) === normalizeText(label))
        ).length,
      })),
    [allLabels, groups]
  );

  useEffect(() => {
    if (selectedManagedLabel && allLabels.some((label) => normalizeText(label) === normalizeText(selectedManagedLabel))) {
      return;
    }
    setSelectedManagedLabel(allLabels[0] || "");
    setRenameLabelValue(allLabels[0] || "");
  }, [allLabels, selectedManagedLabel]);

  useEffect(() => {
    setRenameLabelValue(selectedManagedLabel || "");
  }, [selectedManagedLabel]);

  const visibleGroups = useMemo(() => {
    const query = normalizeText(groupQuery);
    const filtered = groups
      .filter((group) => {
        if (archiveFilter === "active" && group.isArchived) return false;
        if (archiveFilter === "archived" && !group.isArchived) return false;
        if (statusFilter !== "all" && String(group.status || "em_analise") !== statusFilter) return false;
        if (activeLabelFilters.length) {
          const labels = new Set((group.labels || []).map((entry) => normalizeText(entry)));
          if (!activeLabelFilters.every((entry) => labels.has(normalizeText(entry)))) return false;
        }
        if (!query) return true;
        const haystack = [
          group.name,
          group.description,
          statusLabel(group.status),
          ...(group.labels || []),
        ]
          .filter(Boolean)
          .join(" ")
          .toLowerCase();
        return haystack.includes(query);
      });

    const sorted = filtered.sort((a, b) => {
      const favoriteDelta = Number(favoriteGroupSet.has(b.id)) - Number(favoriteGroupSet.has(a.id));
      if (favoriteDelta) return favoriteDelta;
      const aUpdated = String(a.updatedAt || a.createdAt || "");
      const bUpdated = String(b.updatedAt || b.createdAt || "");
      if (aUpdated !== bUpdated) return bUpdated.localeCompare(aUpdated);
      return String(a.name || "").localeCompare(String(b.name || ""), "pt-PT");
    });

    const hasManualFilters = statusFilter !== "all" || archiveFilter !== "active" || activeLabelFilters.length > 0;
    if (query || hasManualFilters) return sorted;
    return sorted.slice(0, 8);
  }, [activeLabelFilters, archiveFilter, favoriteGroupSet, groupQuery, groups, statusFilter]);

  const exactGroupMatch = useMemo(
    () => groups.some((group) => normalizeText(group.name) === normalizeText(groupQuery)),
    [groupQuery, groups]
  );

  const visibleGroupEmails = useMemo(() => {
    const query = normalizeText(groupEmailQuery);
    return [...groupEmails]
      .filter((email) => {
        if (!query) return true;
        return [email.subject, email.fromName, email.fromEmail].filter(Boolean).join(" ").toLowerCase().includes(query);
      })
      .sort((a, b) => String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || "")));
  }, [groupEmailQuery, groupEmails]);

  const visibleLibraryEmails = useMemo(() => {
    const rows = [...libraryEmails];
    const hasCurrent = rows.some((row) => makeEmailKey(row) === currentEmailKey);
    if (
      !hasCurrent
      && (currentEmailPayload.itemId || currentEmailPayload.internetMessageId || currentEmailPayload.conversationId)
      && (!selectedGroupId || !groupEmails.some((row) => makeEmailKey(row) === currentEmailKey))
    ) {
      rows.unshift({
        ...currentEmailPayload,
        relatedGroups: [],
        relatedRecords: [],
      });
    }
    return rows;
  }, [currentEmailKey, currentEmailPayload, groupEmails, libraryEmails, selectedGroupId]);

  const selectedLibraryRows = useMemo(
    () => visibleLibraryEmails.filter((row) => selectedLibraryKeys.includes(makeEmailKey(row))),
    [selectedLibraryKeys, visibleLibraryEmails]
  );

  const selectedGroupEmailRows = useMemo(
    () => visibleGroupEmails.filter((row) => selectedGroupEmailKeys.includes(makeEmailKey(row))),
    [selectedGroupEmailKeys, visibleGroupEmails]
  );

  const currentEmailAlreadyLinked = groupEmails.some((email) => makeEmailKey(email) === currentEmailKey || isCurrentContextEmail(email, ctx));

  async function refreshAll() {
    setReloadToken((value) => value + 1);
  }

  async function persistLabelCatalog(nextCatalog: string[]) {
    await saveSettings({ groupLabelCatalog: mergeLabelCatalog(nextCatalog) });
  }

  async function ensureCatalogLabels(labels: string[]): Promise<string[]> {
    if (!labelsManagerEnabled) return parseLabels(labels);
    const mergedCatalog = mergeLabelCatalog(settings?.groupLabelCatalog || [], labels);
    await persistLabelCatalog(mergedCatalog);
    return canonicalizeLabelsWithCatalog(labels, mergedCatalog);
  }

  async function createCatalogLabel(rawValue: string): Promise<string | null> {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return null;
    }
    const label = String(rawValue || "").trim();
    if (!label) {
      setMsg("Escreve um nome para a nova etiqueta.");
      return null;
    }
    const nextCatalog = mergeLabelCatalog(settings?.groupLabelCatalog || [], [label]);
    await persistLabelCatalog(nextCatalog);
    return nextCatalog.find((entry) => normalizeText(entry) === normalizeText(label)) || label;
  }

  async function handleCreateGroup() {
    const name = String(groupQuery || "").trim();
    if (!name) {
      setMsg("Escreve um nome para pesquisar ou criar o grupo.");
      return;
    }
    setBusy(true);
    try {
      const group = await createLinkGroup({
        name,
        labels: [],
        status: "em_analise",
        documentsEnabled: true,
      });
      setGroupQuery("");
      setSelectedGroupId(group.id);
      setView("detail");
      await refreshAll();
      setMsg("Grupo criado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSaveGroup() {
    if (!selectedGroup) return;
    const name = String(draft.name || "").trim();
    if (!name) {
      setMsg("O grupo precisa de um nome.");
      return;
    }
    setBusy(true);
    try {
      const labels = await ensureCatalogLabels(parseLabels(draft.labelsText));
      await updateLinkGroup(selectedGroup.id, {
        name,
        description: draft.description,
        status: draft.status,
        labels,
        documentsEnabled: draft.documentsEnabled,
        isArchived: draft.isArchived,
      });
      await refreshAll();
      setMsg("Grupo atualizado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleDeleteGroup() {
    if (!selectedGroup) return;
    if (!window.confirm(`Eliminar o grupo "${selectedGroup.name}" e todas as ligacoes?`)) return;
    setBusy(true);
    try {
      await deleteLinkGroup(selectedGroup.id);
      setSelectedGroupId("");
      setView("groups");
      await refreshAll();
      setMsg("Grupo eliminado.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel eliminar o grupo.");
    } finally {
      setBusy(false);
    }
  }

  async function handleAddCurrentEmail() {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      await addEmailToLinkGroup(selectedGroup.id, {
        ...currentEmailPayload,
        membershipKind: linkKind,
      });
      setActiveGroupForCurrentEmail(selectedGroup.id);
      await refreshAll();
      setMsg(`Email atual associado ao grupo como ${membershipKindLabel(linkKind).toLowerCase()}.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel associar o email atual.");
    } finally {
      setBusy(false);
    }
  }

  async function handleAddSelectedLibrary() {
    if (!selectedGroup || !selectedLibraryRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedLibraryRows) {
        await addEmailToLinkGroup(selectedGroup.id, {
          ...email,
          membershipKind: linkKind,
        });
      }
      setSelectedLibraryKeys([]);
      await refreshAll();
      setMsg(`${selectedLibraryRows.length} email(s) associados ao grupo como ${membershipKindLabel(linkKind).toLowerCase()}.`);
      setView("detail");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel associar os emails selecionados.");
    } finally {
      setBusy(false);
    }
  }

  async function handleSetSelectedEmailsKind(nextKind: MembershipKind) {
    if (!selectedGroup || !selectedGroupEmailRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedGroupEmailRows) {
        await addEmailToLinkGroup(selectedGroup.id, {
          ...email,
          membershipKind: nextKind,
        });
      }
      await refreshAll();
      setMsg(`${selectedGroupEmailRows.length} email(s) marcados como ${membershipKindLabel(nextKind).toLowerCase()}.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar o tipo de ligacao.");
    } finally {
      setBusy(false);
    }
  }

  async function handleRemoveSelectedEmails() {
    if (!selectedGroup || !selectedGroupEmailRows.length) return;
    setBusy(true);
    try {
      for (const email of selectedGroupEmailRows) {
        await removeEmailFromLinkGroup(selectedGroup.id, {
          ...email,
          emailKey: makeEmailKey(email),
        });
      }
      setSelectedGroupEmailKeys([]);
      await refreshAll();
      setMsg(`${selectedGroupEmailRows.length} email(s) removidos do grupo.`);
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel remover os emails selecionados.");
    } finally {
      setBusy(false);
    }
  }

  function toggleLabelFilter(label: string) {
    setActiveLabelFilters((current) => {
      const key = normalizeText(label);
      return current.some((entry) => normalizeText(entry) === key)
        ? current.filter((entry) => normalizeText(entry) !== key)
        : [...current, label];
    });
  }

  async function toggleFavoriteGroup(groupId: string) {
    const next = favoriteGroupSet.has(groupId)
      ? favoriteGroupIds.filter((id) => id !== groupId)
      : [groupId, ...favoriteGroupIds.filter((id) => id !== groupId)];
    try {
      await saveSettings({ groupFavoriteIds: next });
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel atualizar os favoritos.");
    }
  }

  async function handleCreateCatalogLabel() {
    setBusy(true);
    try {
      const canonical = await createCatalogLabel(newCatalogLabel);
      if (!canonical) return;
      setNewCatalogLabel("");
      setSelectedManagedLabel(canonical);
      setRenameLabelValue(canonical);
      setMsg("Etiqueta adicionada ao catalogo.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel criar a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function handleRenameManagedLabel() {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return;
    }
    const source = String(selectedManagedLabel || "").trim();
    const target = String(renameLabelValue || "").trim();
    if (!source) {
      setMsg("Seleciona primeiro uma etiqueta.");
      return;
    }
    if (!target) {
      setMsg("Escreve o novo nome da etiqueta.");
      return;
    }
    setBusy(true);
    try {
      const nextCatalog = mergeLabelCatalog(
        (settings?.groupLabelCatalog || []).map((entry) => (normalizeText(entry) === normalizeText(source) ? target : entry)),
        [target]
      );
      for (const group of groups) {
        const labels = group.labels || [];
        if (!labels.some((entry) => normalizeText(entry) === normalizeText(source))) continue;
        const nextLabels = canonicalizeLabelsWithCatalog(
          labels.map((entry) => (normalizeText(entry) === normalizeText(source) ? target : entry)),
          nextCatalog
        );
        await updateLinkGroup(group.id, { labels: nextLabels });
      }
      await persistLabelCatalog(nextCatalog);
      setSelectedManagedLabel(target);
      setRenameLabelValue(target);
      await refreshAll();
      setMsg("Etiqueta renomeada em todos os grupos.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel renomear a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function handleDeleteManagedLabel() {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return;
    }
    const label = String(selectedManagedLabel || "").trim();
    if (!label) {
      setMsg("Seleciona primeiro uma etiqueta.");
      return;
    }
    if (!window.confirm(`Eliminar a etiqueta "${label}" de todos os grupos?`)) return;
    setBusy(true);
    try {
      for (const group of groups) {
        const labels = group.labels || [];
        if (!labels.some((entry) => normalizeText(entry) === normalizeText(label))) continue;
        await updateLinkGroup(group.id, {
          labels: labels.filter((entry) => normalizeText(entry) !== normalizeText(label)),
        });
      }
      const nextCatalog = (settings?.groupLabelCatalog || []).filter((entry) => normalizeText(entry) !== normalizeText(label));
      await persistLabelCatalog(nextCatalog);
      setSelectedManagedLabel("");
      setRenameLabelValue("");
      await refreshAll();
      setMsg("Etiqueta eliminada do catalogo e dos grupos.");
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel eliminar a etiqueta.");
    } finally {
      setBusy(false);
    }
  }

  async function openExplorer() {
    if (!selectedGroup) return;
    try {
      await openGroupExplorer({ groupId: selectedGroup.id });
    } catch (error: any) {
      setMsg(error?.message || "Nao foi possivel abrir o explorer do grupo.");
    }
  }

  const trackTransform = view === "groups" ? "translateX(0%)" : view === "detail" ? "translateX(-33.3333%)" : "translateX(-66.6667%)";
  const configView = view === "settings" || view === "labels";

  const headerTitle = view === "settings"
    ? "Settings dos grupos"
    : view === "labels"
      ? "Gestor de etiquetas"
      : "Grupos";

  const headerHint = view === "settings"
    ? "Configuracoes locais."
    : view === "labels"
      ? "Catalogo central."
      : "Ligacoes manuais.";

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.title}>{headerTitle}</div>
          <div style={S.headerHint}>{headerHint}</div>
        </div>
        <div style={S.headerActions}>
          {configView ? (
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => setView(view === "labels" ? "settings" : "groups")}
              disabled={busy}
            >
              Voltar
            </button>
          ) : (
            <>
              <button type="button" style={S.secondaryBtn} onClick={() => void refreshAll()} disabled={groupsLoading || busy}>
                <Icons.RefreshCw size={12} />
                Atualizar
              </button>
              <button
                type="button"
                style={S.iconGearBtn}
                onClick={() => setView("settings")}
                disabled={busy}
                title="Settings dos grupos"
              >
                <Icons.Settings size={14} />
              </button>
            </>
          )}
        </div>
      </div>

      <div style={S.viewport}>
        {view === "settings" ? (
          <section style={S.cleanPanel}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Settings dos grupos</div>
                  <HelpHint text="Area limpa para extras do modulo Grupos. Aqui entras nos ecras de configuracao sem contexto do email aberto." title="Ajuda: Settings dos grupos" />
                </div>
              </div>
            </div>

            <div style={S.settingMenuGrid}>
              <button type="button" style={S.settingEntry} onClick={() => setView("labels")}>
                <div style={S.settingEntryBody}>
                  <div style={S.settingEntryTitle}>Etiquetas</div>
                </div>
                <span style={S.settingEntryMeta}>{labelsManagerEnabled ? "Ativo" : "Desativado"}</span>
              </button>

              <div style={S.card}>
                <div style={S.fieldLabel}>Ativacao global</div>
                <div style={S.smallMeta}>A ativacao continua em Settings {" > "} Grupos.</div>
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings gerais
                  </button>
                </div>
              </div>
            </div>
          </section>
        ) : view === "labels" ? (
          <section style={S.cleanPanel}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Gestor de etiquetas</div>
                  <HelpHint text="Catalogo central de etiquetas. Cria, renomeia e remove etiquetas sem depender do email ou do grupo ativo." title="Ajuda: Gestor de etiquetas" />
                </div>
              </div>
            </div>

            {!labelsManagerEnabled ? (
              <div style={S.card}>
                <PanelState
                  compact
                  tone="info"
                  title="Gestor de etiquetas desativado"
                  description="Ativa a funcionalidade em Settings > Grupos para usar o catalogo central."
                />
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings gerais
                  </button>
                </div>
              </div>
            ) : (
              <div style={S.settingsGrid}>
                <div style={S.card}>
                  <div style={S.fieldLabel}>Nova etiqueta</div>
                  <div style={S.inlineRow}>
                    <input
                      style={S.input}
                      value={newCatalogLabel}
                      onChange={(event) => setNewCatalogLabel(event.target.value)}
                      placeholder="Ex.: marca, cliente, ganho"
                    />
                    <button type="button" style={S.primaryBtn} onClick={() => void handleCreateCatalogLabel()} disabled={busy}>
                      <Icons.Plus size={12} />
                      Criar
                    </button>
                  </div>
                  <div style={S.smallMeta}>Catalogo central usado para evitar duplicados e normalizar etiquetas dos grupos.</div>
                </div>

                <div style={S.settingsColumns}>
                  <div style={S.managerList}>
                    <div style={S.fieldLabel}>Etiquetas existentes</div>
                    {!labelUsage.length ? (
                      <PanelState compact tone="info" title="Sem etiquetas" description="Cria a primeira etiqueta ou guarda etiquetas num grupo." />
                    ) : (
                      labelUsage.map((entry) => {
                        const active = normalizeText(entry.label) === normalizeText(selectedManagedLabel);
                        return (
                          <button
                            key={entry.label}
                            type="button"
                            style={active ? S.managerRowActive : S.managerRow}
                            onClick={() => setSelectedManagedLabel(entry.label)}
                          >
                            <span>{entry.label}</span>
                            <span style={S.managerCount}>{entry.count}</span>
                          </button>
                        );
                      })
                    )}
                  </div>

                  <div style={S.card}>
                    <div style={S.fieldLabel}>Etiqueta selecionada</div>
                    {selectedManagedLabel ? (
                      <>
                        <input
                          style={S.input}
                          value={renameLabelValue}
                          onChange={(event) => setRenameLabelValue(event.target.value)}
                          placeholder="Novo nome da etiqueta"
                        />
                        <div style={S.smallMeta}>As acoes abaixo atualizam todos os grupos que usam esta etiqueta.</div>
                        <div style={S.inlineRow}>
                          <button type="button" style={S.primaryBtn} onClick={() => void handleRenameManagedLabel()} disabled={busy}>
                            <Icons.Save size={12} />
                            Renomear globalmente
                          </button>
                          <button type="button" style={S.dangerBtn} onClick={() => void handleDeleteManagedLabel()} disabled={busy}>
                            <Icons.Trash size={12} />
                            Eliminar globalmente
                          </button>
                        </div>
                      </>
                    ) : (
                      <PanelState compact tone="info" title="Seleciona uma etiqueta" description="Escolhe uma etiqueta da lista para a renomear ou eliminar." />
                    )}
                  </div>
                </div>
              </div>
            )}
          </section>
        ) : (
          <div style={{ ...S.track, transform: trackTransform }}>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Grupos</div>
                  <HelpHint text="Seleciona um grupo para o gerir. Usa os filtros para reduzir a lista e abre o detalhe so quando precisares." title="Ajuda: Grupos" />
                </div>
              </div>
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Grupo</div>
                <HelpHint text="Pesquisa um grupo pelo nome. Se nao existir nenhum igual, o mesmo campo permite criar logo esse grupo." title="Ajuda: Grupo" />
              </div>
              <div style={S.labelSearchRow}>
                <input
                  style={S.input}
                  value={groupQuery}
                  onChange={(event) => setGroupQuery(event.target.value)}
                  placeholder="Pesquisar ou criar grupo"
                />
                <button
                  type="button"
                  style={!exactGroupMatch && String(groupQuery || "").trim() ? S.primaryBtn : S.iconGhostBtn}
                  onClick={() => void handleCreateGroup()}
                  disabled={busy || !String(groupQuery || "").trim() || exactGroupMatch}
                  title={!exactGroupMatch && String(groupQuery || "").trim() ? `Criar "${String(groupQuery || "").trim()}"` : "Pesquisar grupos existentes"}
                >
                  {!exactGroupMatch && String(groupQuery || "").trim() ? <Icons.Plus size={12} /> : <Icons.Search size={12} />}
                </button>
              </div>
              {!exactGroupMatch && String(groupQuery || "").trim() ? (
                <div style={S.smallMeta}>Criar grupo: <b>{String(groupQuery || "").trim()}</b></div>
              ) : (
                <div style={S.smallMeta}>{favoriteGroupIds.length ? "Favoritos e recentes" : "Recentes"}</div>
              )}
            </div>

            <div style={S.card}>
              <div style={S.sectionTitleRow}>
                <div style={S.fieldLabel}>Filtros</div>
                <HelpHint text="Filtra por texto, estado, arquivo e etiquetas. A explicacao fica aqui para nao sobrecarregar o ecran." title="Ajuda: Filtros" />
              </div>
              <div style={S.inlineRow}>
                <select style={S.select} value={statusFilter} onChange={(event) => setStatusFilter(event.target.value as GroupStatusFilter)}>
                  <option value="all">Todos os estados</option>
                  {STATUS_OPTIONS.map((option) => (
                    <option key={option.value} value={option.value}>{option.label}</option>
                  ))}
                </select>
                <select style={S.select} value={archiveFilter} onChange={(event) => setArchiveFilter(event.target.value as GroupArchiveFilter)}>
                  <option value="active">Ativos</option>
                  <option value="archived">Arquivados</option>
                  <option value="all">Todos</option>
                </select>
              </div>
              {allLabels.length ? (
                <div style={S.chipWrap}>
                  {allLabels.map((label) => {
                    const active = activeLabelFilters.some((entry) => normalizeText(entry) === normalizeText(label));
                    return (
                      <button key={label} type="button" style={active ? S.chipActive : S.chip} onClick={() => toggleLabelFilter(label)}>
                        {label}
                      </button>
                    );
                  })}
                </div>
              ) : null}
            </div>

            {groupsError ? <PanelState compact tone="error" title="Falha a carregar grupos" description={groupsError} /> : null}
            {groupsLoading && !groups.length ? <PanelState compact tone="loading" title="A carregar grupos" description="A preparar a lista central de grupos." /> : null}
            {!groupsLoading && !visibleGroups.length ? <PanelState compact tone="info" title="Sem grupos visiveis" description="Cria um grupo novo ou alarga os filtros." /> : null}

            <div style={S.listWrap}>
              {visibleGroups.map((group) => {
                const selected = group.id === selectedGroupId;
                const favorite = favoriteGroupSet.has(group.id);
                return (
                  <div key={group.id} style={selected ? S.groupItemActive : S.groupItem}>
                    <div style={S.groupItemHead}>
                      <button
                        type="button"
                        style={S.groupMainBtn}
                        onClick={() => {
                          setSelectedGroupId(group.id);
                          setView("detail");
                        }}
                      >
                        <div style={S.groupName}>{group.name}</div>
                        <span style={{ ...S.statusBadge, ...(group.status === "concluido" ? S.statusDone : group.status === "em_progresso" ? S.statusProgress : S.statusAnalysis) }}>
                          {statusLabel(group.status)}
                        </span>
                      </button>
                      <button
                        type="button"
                        style={favorite ? S.favoriteBtnActive : S.favoriteBtn}
                        onClick={() => void toggleFavoriteGroup(group.id)}
                        title={favorite ? "Remover dos favoritos" : "Adicionar aos favoritos"}
                      >
                        <Icons.Star size={11} />
                      </button>
                    </div>
                    {group.description ? <div style={S.groupDescription}>{group.description}</div> : null}
                    <div style={S.groupMeta}>
                      <span>{group.memberCount || 0} email(s)</span>
                      {favorite ? <span>Favorito</span> : null}
                      {group.isArchived ? <span>Arquivado</span> : null}
                    </div>
                    {group.labels?.length ? (
                      <div style={S.groupLabels}>
                        {group.labels.slice(0, 2).map((label) => (
                          <span key={label} style={S.labelBadge}>{label}</span>
                        ))}
                      </div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          </section>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("groups")}>Voltar</button>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>{selectedGroup ? selectedGroup.name : "Detalhe do grupo"}</div>
                  <HelpHint text="Fluxo rapido: escolhe Principal ou Referencia e liga o email atual. A biblioteca fica para outros emails ja registados." title="Ajuda: Detalhe do grupo" />
                </div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Seleciona um grupo" description="Escolhe um grupo na lista para o gerir." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.sectionTitleRow}>
                    <div style={S.fieldLabel}>Dados</div>
                    <HelpHint text="Nome, estado, arquivo e etiquetas do grupo." title="Ajuda: Dados do grupo" />
                  </div>
                  <input style={S.input} value={draft.name} onChange={(event) => setDraft((current) => ({ ...current, name: event.target.value }))} placeholder="Nome" />
                  <textarea style={S.textarea} value={draft.description} onChange={(event) => setDraft((current) => ({ ...current, description: event.target.value }))} placeholder="Descricao do grupo" />
                  <div style={S.inlineRow}>
                    <select style={S.select} value={draft.status} onChange={(event) => setDraft((current) => ({ ...current, status: event.target.value as GroupDraft["status"] }))}>
                      {STATUS_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                    <label style={S.toggleRow}>
                      <input type="checkbox" checked={draft.isArchived} onChange={(event) => setDraft((current) => ({ ...current, isArchived: event.target.checked }))} />
                      <span>Arquivado</span>
                    </label>
                  </div>
                  <LabelPicker
                    value={draft.labelsText}
                    onChange={(next) => setDraft((current) => ({ ...current, labelsText: next }))}
                    catalog={allLabels}
                    enabled={labelsManagerEnabled}
                    busy={busy}
                    placeholder="Pesquisar etiqueta"
                    helpText="Pesquisa e adiciona etiquetas ao grupo. Se nao existir nenhuma, o botao cria logo a etiqueta no catalogo."
                    onCreateLabel={createCatalogLabel}
                  />
                  <label style={S.toggleRow}>
                    <input type="checkbox" checked={draft.documentsEnabled} onChange={(event) => setDraft((current) => ({ ...current, documentsEnabled: event.target.checked }))} />
                    <span>Documentos ativos neste grupo</span>
                  </label>
                  <div style={S.inlineRow}>
                    <button type="button" style={S.primaryBtn} onClick={() => void handleSaveGroup()} disabled={busy || !draftChanged(selectedGroup, draft)}>
                      <Icons.Save size={12} />
                      Guardar
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void openExplorer()} disabled={busy}>
                      <Icons.ExternalLink size={12} />
                      Explorer
                    </button>
                    <button type="button" style={S.dangerBtn} onClick={() => void handleDeleteGroup()} disabled={busy}>
                      <Icons.Trash size={12} />
                      Eliminar
                    </button>
                  </div>
                </div>

                <div style={S.card}>
                  <div style={S.sectionRow}>
                    <div>
                      <div style={S.sectionTitleRow}>
                        <div style={S.fieldLabel}>Emails</div>
                        <HelpHint text="Aqui ligas o email atual ao grupo ou mudas a ligacao dos emails ja guardados entre Principal e Referencia." title="Ajuda: Emails do grupo" />
                      </div>
                      <div style={S.smallMeta}>
                        {groupEmails.length} email(s) ligados
                        {currentEmailAlreadyLinked ? " · email atual ligado" : ""}
                      </div>
                    </div>
                    <div style={S.inlineActions}>
                      <div style={S.togglePillBar} title="Tipo de ligacao">
                        {MEMBERSHIP_OPTIONS.map((option) => (
                          <button
                            key={option.value}
                            type="button"
                            style={linkKind === option.value ? S.togglePillActive : S.togglePill}
                            onClick={() => setLinkKind(option.value)}
                          >
                            {option.label}
                          </button>
                        ))}
                      </div>
                      <button type="button" style={S.secondaryBtn} onClick={() => void handleAddCurrentEmail()} disabled={busy || currentEmailAlreadyLinked} title="Ligar o email aberto ao grupo selecionado">
                        <Icons.Link size={12} />
                        Ligar atual
                      </button>
                      <button type="button" style={S.secondaryBtn} onClick={() => setView("library")} disabled={busy} title="Abrir biblioteca de emails ja registados">
                        <Icons.Plus size={12} />
                        Biblioteca
                      </button>
                    </div>
                  </div>
                  <div style={S.inlineRow}>
                    <input style={S.input} value={groupEmailQuery} onChange={(event) => setGroupEmailQuery(event.target.value)} placeholder="Filtrar emails do grupo" />
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleSetSelectedEmailsKind("principal")} disabled={busy || !selectedGroupEmailRows.length}>
                      Tornar principal
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleSetSelectedEmailsKind("referencia")} disabled={busy || !selectedGroupEmailRows.length}>
                      Tornar referencia
                    </button>
                    <button type="button" style={S.secondaryBtn} onClick={() => void handleRemoveSelectedEmails()} disabled={busy || !selectedGroupEmailRows.length}>
                      Remover selecionados
                    </button>
                  </div>

                  {groupEmailsLoading ? <PanelState compact tone="loading" title="A carregar emails" description="A ler os emails associados a este grupo." /> : null}
                  {!groupEmailsLoading && !visibleGroupEmails.length ? <PanelState compact tone="info" title="Sem emails ligados" description="Usa 'Ligar atual' ou abre a biblioteca." /> : null}

                  <div style={S.emailList}>
                    {visibleGroupEmails.map((email) => {
                      const rowKey = makeEmailKey(email);
                      const selected = selectedGroupEmailKeys.includes(rowKey);
                      const isCurrent = rowKey === currentEmailKey || isCurrentContextEmail(email, ctx);
                      return (
                        <div key={rowKey} style={selected ? S.emailRowActive : S.emailRow}>
                          <label style={S.checkboxCell}>
                            <input
                              type="checkbox"
                              checked={selected}
                              onChange={(event) => {
                                setSelectedGroupEmailKeys((current) => event.target.checked ? [...current, rowKey] : current.filter((entry) => entry !== rowKey));
                              }}
                            />
                          </label>
                          <button type="button" style={S.emailMain} onClick={() => openLinkedOutlookEmail(email as any)}>
                            <div style={S.emailSubject}>
                              {email.subject || "Sem assunto"}
                              <span style={email.membershipKind === "referencia" ? S.refBadge : S.principalBadge}>
                                {membershipKindLabel(email.membershipKind)}
                              </span>
                              {isCurrent ? <span style={S.currentEmailBadge}>Atual</span> : null}
                            </div>
                            <div style={S.emailMeta}>{email.fromName || email.fromEmail || "Sem remetente"} · {formatDate(email.messageDateIso || email.receivedAtIso)}</div>
                          </button>
                          <button type="button" style={S.iconOnlyBtn} onClick={() => openLinkedOutlookEmail(email as any)} title="Abrir email">
                            <Icons.ExternalLink size={12} />
                          </button>
                        </div>
                      );
                    })}
                  </div>
                </div>
              </>
            )}
          </section>

          <section style={S.panel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("detail")}>Voltar</button>
              <div>
                <div style={S.sectionTitleRow}>
                  <div style={S.panelTitle}>Biblioteca</div>
                  <HelpHint text="Usa esta vista apenas quando queres ligar outros emails ja registados. Para o email atual, usa a ligacao rapida no detalhe." title="Ajuda: Biblioteca" />
                </div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Sem grupo selecionado" description="Escolhe primeiro um grupo." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.sectionTitleRow}>
                    <div style={S.fieldLabel}>Pesquisa</div>
                    <HelpHint text="Pesquisa por assunto, remetente, grupo ou registo e liga os emails selecionados ao grupo atual." title="Ajuda: Pesquisa de emails" />
                  </div>
                  <div style={S.inlineRow}>
                    <input style={S.input} value={libraryQuery} onChange={(event) => setLibraryQuery(event.target.value)} placeholder="Assunto, remetente, grupo, registo..." />
                    <div style={S.togglePillBar} title="Tipo de ligacao">
                      {MEMBERSHIP_OPTIONS.map((option) => (
                        <button
                          key={option.value}
                          type="button"
                          style={linkKind === option.value ? S.togglePillActive : S.togglePill}
                          onClick={() => setLinkKind(option.value)}
                        >
                          {option.label}
                        </button>
                      ))}
                    </div>
                    <button type="button" style={S.primaryBtn} onClick={() => void handleAddSelectedLibrary()} disabled={busy || !selectedLibraryRows.length}>
                      <Icons.Link size={12} />
                      Ligar
                    </button>
                  </div>
                  <div style={S.smallMeta}>Grupo alvo: {selectedGroup.name}</div>
                </div>

                {libraryLoading ? <PanelState compact tone="loading" title="A pesquisar emails" description="A varrer os emails ja registados pelo cockpit." /> : null}
                {!libraryLoading && !visibleLibraryEmails.length ? <PanelState compact tone="info" title="Sem resultados" description="Alarga a pesquisa ou usa o email atual." /> : null}

                <div style={S.emailList}>
                  {visibleLibraryEmails.map((email) => {
                    const rowKey = makeEmailKey(email);
                    const selected = selectedLibraryKeys.includes(rowKey);
                    const isCurrent = rowKey === currentEmailKey;
                    return (
                      <div key={rowKey} style={selected ? S.emailRowActive : S.emailRow}>
                        <label style={S.checkboxCell}>
                          <input
                            type="checkbox"
                            checked={selected}
                            onChange={(event) => {
                              setSelectedLibraryKeys((current) => event.target.checked ? [...current, rowKey] : current.filter((entry) => entry !== rowKey));
                            }}
                          />
                        </label>
                        <button type="button" style={S.emailMain} onClick={() => openLinkedOutlookEmail(email as any)}>
                          <div style={S.emailSubject}>
                            {email.subject || "Sem assunto"}
                            {isCurrent ? <span style={S.currentEmailBadge}>Atual</span> : null}
                          </div>
                          <div style={S.emailMeta}>{email.fromName || email.fromEmail || "Sem remetente"} · {formatDate(email.messageDateIso || email.receivedAtIso)}</div>
                          {(email.relatedGroups?.length || email.relatedRecords?.length) ? (
                            <div style={S.emailAssociations}>
                              {(email.relatedGroups || []).slice(0, 3).map((group) => (
                                <span key={`${rowKey}:g:${group.id}`} style={group.relationKind === "referencia" ? S.refBadge : S.labelBadge}>
                                  {group.name || group.id}
                                  {group.relationKind === "referencia" ? " · Ref" : ""}
                                </span>
                              ))}
                              {(email.relatedRecords || []).slice(0, 2).map((record) => (
                                <span key={`${rowKey}:r:${record.model}:${record.recordId}`} style={S.recordBadge}>{record.recordName || `${record.model}#${record.recordId}`}</span>
                              ))}
                            </div>
                          ) : null}
                        </button>
                        <button type="button" style={S.iconOnlyBtn} onClick={() => openLinkedOutlookEmail(email as any)} title="Abrir email">
                          <Icons.ExternalLink size={12} />
                        </button>
                      </div>
                    );
                  })}
                </div>
              </>
            )}
          </section>
          </div>
        )}
      </div>
    </div>
  );
};

const baseButton: React.CSSProperties = {
  borderRadius: 10,
  border: "1px solid var(--iccc-card-border)",
  padding: "7px 10px",
  fontSize: 11,
  fontWeight: 600,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 6,
};

const S: Record<string, React.CSSProperties> = {
  root: { display: "grid", gap: 8, minHeight: "100%", position: "relative" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: 10, borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  headerActions: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  kicker: { fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.05em" },
  title: { fontSize: 15, fontWeight: 700, color: "var(--iccc-text)" },
  headerHint: { fontSize: 11, color: "var(--iccc-text-muted)", marginTop: 1 },
  settingsGrid: { display: "grid", gap: 10 },
  settingsColumns: { display: "grid", gridTemplateColumns: "minmax(210px, 250px) minmax(0, 1fr)", gap: 10, alignItems: "start" },
  managerList: { display: "grid", gap: 6, padding: 10, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", alignContent: "start" },
  managerRow: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "#fff", padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 12, fontWeight: 600 },
  managerRowActive: { width: "100%", borderRadius: 10, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: "8px 10px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 12, fontWeight: 600 },
  managerCount: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 22, height: 22, borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 700 },
  viewport: { overflow: "hidden", borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  track: { width: "300%", display: "flex", transition: "transform 0.22s ease" },
  panel: { width: "33.3333%", padding: 10, display: "grid", alignContent: "start", gap: 10, minHeight: "calc(100vh - 220px)", boxSizing: "border-box" },
  cleanPanel: { padding: 10, display: "grid", alignContent: "start", gap: 10, minHeight: "calc(100vh - 220px)", boxSizing: "border-box" },
  settingMenuGrid: { display: "grid", gap: 10, alignContent: "start" },
  settingEntry: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.78)", padding: 10, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer" },
  settingEntryBody: { display: "grid", gap: 0, minWidth: 0 },
  settingEntryTitle: { fontSize: 13, fontWeight: 700, color: "var(--iccc-text)" },
  settingEntryMeta: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 10, fontWeight: 700, whiteSpace: "nowrap" },
  panelHeader: { display: "flex", alignItems: "center", gap: 8 },
  panelTitle: { fontSize: 14, fontWeight: 700, color: "var(--iccc-text)" },
  sectionTitleRow: { display: "inline-flex", alignItems: "center", gap: 6, flexWrap: "wrap" },
  backBtn: { ...baseButton, background: "transparent", color: "var(--iccc-text)" },
  primaryBtn: { ...baseButton, background: "linear-gradient(180deg, rgba(96, 165, 250, 0.95) 0%, rgba(37, 99, 235, 0.95) 100%)", color: "#fff", border: "1px solid rgba(37, 99, 235, 0.35)" },
  secondaryBtn: { ...baseButton, background: "rgba(255,255,255,0.78)", color: "var(--iccc-text)" },
  iconGearBtn: { ...baseButton, width: 38, height: 38, padding: 0, background: "rgba(255,255,255,0.78)", color: "var(--iccc-text)" },
  iconGhostBtn: { ...baseButton, width: 34, height: 34, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
  dangerBtn: { ...baseButton, background: "rgba(254, 226, 226, 0.95)", color: "#b91c1c", border: "1px solid rgba(239, 68, 68, 0.25)" },
  card: { display: "grid", gap: 6, padding: 10, borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)" },
  fieldLabel: { fontSize: 10, fontWeight: 700, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.05em" },
  input: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)", boxSizing: "border-box" },
  compactSelect: { borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "7px 9px", background: "#fff", fontSize: 11, color: "var(--iccc-text)", minWidth: 120 },
  textarea: { width: "100%", minHeight: 72, borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)", boxSizing: "border-box", resize: "vertical" },
  select: { width: "100%", borderRadius: 10, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)" },
  inlineRow: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  toggleRow: { display: "inline-flex", gap: 6, alignItems: "center", fontSize: 11, color: "var(--iccc-text)" },
  chipWrap: { display: "flex", flexWrap: "wrap", gap: 8 },
  chip: { ...baseButton, padding: "4px 8px", fontSize: 10, background: "rgba(255,255,255,0.85)", color: "var(--iccc-text)" },
  chipActive: { ...baseButton, padding: "4px 8px", fontSize: 10, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  listWrap: { display: "grid", gap: 6 },
  groupItem: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", padding: 10, display: "grid", gap: 6, cursor: "pointer" },
  groupItemActive: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 10, display: "grid", gap: 6, cursor: "pointer" },
  groupItemHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  groupMainBtn: { border: "none", background: "transparent", padding: 0, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, width: "100%", minWidth: 0, cursor: "pointer", textAlign: "left" },
  groupName: { fontSize: 13, fontWeight: 700, color: "var(--iccc-text)" },
  groupDescription: { fontSize: 11, color: "var(--iccc-text-muted)", lineHeight: 1.35, display: "-webkit-box", WebkitLineClamp: 2, WebkitBoxOrient: "vertical", overflow: "hidden" },
  groupMeta: { display: "flex", gap: 8, flexWrap: "wrap", fontSize: 10, color: "var(--iccc-text-muted)" },
  groupLabels: { display: "flex", gap: 6, flexWrap: "wrap" },
  favoriteBtn: { ...baseButton, width: 28, height: 28, padding: 0, background: "rgba(255,255,255,0.88)", color: "var(--iccc-text-muted)" },
  favoriteBtnActive: { ...baseButton, width: 28, height: 28, padding: 0, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  labelBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 10, fontWeight: 600 },
  principalBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  refBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  recordBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 600 },
  statusBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, border: "1px solid transparent", fontSize: 9, fontWeight: 700, textTransform: "uppercase", letterSpacing: "0.04em" },
  statusAnalysis: { background: "rgba(217, 119, 6, 0.12)", color: "#b45309", borderColor: "rgba(217, 119, 6, 0.2)" },
  statusProgress: { background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", borderColor: "rgba(37, 99, 235, 0.2)" },
  statusDone: { background: "rgba(16, 185, 129, 0.12)", color: "#047857", borderColor: "rgba(16, 185, 129, 0.25)" },
  sectionRow: { display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" },
  smallMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  inlineActions: { display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  emailList: { display: "grid", gap: 6 },
  emailRow: { display: "grid", gridTemplateColumns: "22px minmax(0, 1fr) 30px", gap: 6, alignItems: "center", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.82)", padding: 8 },
  emailRowActive: { display: "grid", gridTemplateColumns: "22px minmax(0, 1fr) 30px", gap: 6, alignItems: "center", borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 8 },
  checkboxCell: { display: "inline-flex", alignItems: "center", justifyContent: "center" },
  emailMain: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 4, cursor: "pointer", minWidth: 0 },
  emailSubject: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap", fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" },
  emailMeta: { fontSize: 10, color: "var(--iccc-text-muted)" },
  emailAssociations: { display: "flex", gap: 6, flexWrap: "wrap" },
  currentEmailBadge: { display: "inline-flex", alignItems: "center", padding: "2px 6px", borderRadius: 999, background: "rgba(16, 185, 129, 0.12)", color: "#047857", fontSize: 9, fontWeight: 700, textTransform: "uppercase" },
  iconOnlyBtn: { ...baseButton, width: 30, height: 30, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
  labelPicker: { display: "grid", gap: 5 },
  labelSearchRow: { display: "grid", gridTemplateColumns: "minmax(0, 1fr) 34px", gap: 6, alignItems: "center" },
  labelSuggestionList: { display: "grid", gap: 4, padding: 6, borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)" },
  labelSuggestion: { border: "none", background: "transparent", textAlign: "left", padding: "5px 6px", borderRadius: 8, fontSize: 11, color: "var(--iccc-text)", cursor: "pointer" },
  selectedLabelRow: { display: "flex", flexWrap: "wrap", gap: 6 },
  selectedLabelChip: { ...baseButton, padding: "3px 8px", fontSize: 10, fontWeight: 600, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8" },
  selectedLabelRemove: { fontSize: 12, lineHeight: 1, opacity: 0.8 },
  togglePillBar: { display: "inline-flex", alignItems: "center", gap: 4, padding: 3, borderRadius: 999, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.92)" },
  togglePill: { border: "none", background: "transparent", color: "var(--iccc-text-muted)", padding: "5px 8px", borderRadius: 999, fontSize: 10, fontWeight: 700, cursor: "pointer" },
  togglePillActive: { border: "none", background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", padding: "5px 8px", borderRadius: 999, fontSize: 10, fontWeight: 700, cursor: "pointer" },
};

export default GroupManagerCockpit;
