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
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type GroupManagerView = "groups" | "detail" | "library";
type GroupStatusFilter = "all" | "em_analise" | "em_progresso" | "concluido";
type GroupArchiveFilter = "active" | "archived" | "all";
type MembershipKind = "principal" | "referencia";
type GroupSettingsSection = "labels";

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

export const GroupManagerCockpit: React.FC = () => {
  const { ctx, bodyText, bodyHtml, attachments, setMsg, setActiveGroupForCurrentEmail, settings, openSettingsSection } = useCockpit();
  const [view, setView] = useState<GroupManagerView>("groups");
  const [settingsOpen, setSettingsOpen] = useState(false);
  const [settingsSection, setSettingsSection] = useState<GroupSettingsSection>("labels");
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [groupsLoading, setGroupsLoading] = useState(false);
  const [groupsError, setGroupsError] = useState("");
  const [groupQuery, setGroupQuery] = useState("");
  const [statusFilter, setStatusFilter] = useState<GroupStatusFilter>("all");
  const [archiveFilter, setArchiveFilter] = useState<GroupArchiveFilter>("active");
  const [activeLabelFilters, setActiveLabelFilters] = useState<string[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState("");
  const [draft, setDraft] = useState<GroupDraft>(createDraft(null));
  const [newGroupName, setNewGroupName] = useState("");
  const [newGroupLabels, setNewGroupLabels] = useState("");
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
    return groups
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
      })
      .sort((a, b) => String(a.name || "").localeCompare(String(b.name || ""), "pt-PT"));
  }, [activeLabelFilters, archiveFilter, groupQuery, groups, statusFilter]);

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

  async function handleCreateGroup() {
    const name = String(newGroupName || "").trim();
    if (!name) {
      setMsg("Escreve um nome para criar o grupo.");
      return;
    }
    setBusy(true);
    try {
      const labels = await ensureCatalogLabels(parseLabels(newGroupLabels));
      const group = await createLinkGroup({
        name,
        labels,
        status: "em_analise",
        documentsEnabled: true,
      });
      setNewGroupName("");
      setNewGroupLabels("");
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

  function toggleLabelsTextValue(value: string, label: string): string {
    const current = parseLabels(value);
    const key = normalizeText(label);
    const next = current.some((entry) => normalizeText(entry) === key)
      ? current.filter((entry) => normalizeText(entry) !== key)
      : [...current, label];
    return labelsToText(canonicalizeLabelsWithCatalog(next, settings?.groupLabelCatalog || []));
  }

  async function handleCreateCatalogLabel() {
    if (!labelsManagerEnabled) {
      setMsg("Ativa primeiro o gestor de etiquetas em Settings > Grupos.");
      return;
    }
    const label = String(newCatalogLabel || "").trim();
    if (!label) {
      setMsg("Escreve um nome para a nova etiqueta.");
      return;
    }
    setBusy(true);
    try {
      const nextCatalog = mergeLabelCatalog(settings?.groupLabelCatalog || [], [label]);
      await persistLabelCatalog(nextCatalog);
      const canonical = nextCatalog.find((entry) => normalizeText(entry) === normalizeText(label)) || label;
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

  return (
    <div style={S.root}>
      <div style={S.header}>
        <div>
          <div style={S.kicker}>Gestor de Grupos</div>
          <div style={S.title}>Gestao dedicada de grupos, emails e etiquetas</div>
        </div>
        <div style={S.headerActions}>
          <button type="button" style={S.secondaryBtn} onClick={() => void refreshAll()} disabled={groupsLoading || busy}>
            <Icons.RefreshCw size={12} />
            Atualizar
          </button>
          <button
            type="button"
            style={settingsOpen ? S.secondaryBtnActive : S.secondaryBtn}
            onClick={() => setSettingsOpen((current) => !current)}
            disabled={busy}
            title="Configurar grupos"
          >
            <Icons.Settings size={13} />
            {settingsOpen ? "Fechar" : "Definicoes"}
          </button>
        </div>
      </div>

      {settingsOpen ? (
        <div style={S.settingsShell}>
          <div style={S.settingsMenu}>
            <button
              type="button"
              style={settingsSection === "labels" ? S.settingsMenuItemActive : S.settingsMenuItem}
              onClick={() => setSettingsSection("labels")}
            >
              <Icons.Settings size={13} />
              Gestor de etiquetas
            </button>
          </div>
          <div style={S.settingsContent}>
            <div style={S.settingsContentHead}>
              <div>
                <div style={S.panelTitle}>Definicoes do modulo Grupos</div>
                <div style={S.panelHint}>A operacao diaria das etiquetas fica aqui; a ativacao global continua em Settings.</div>
              </div>
            </div>
            {!labelsManagerEnabled ? (
              <div style={S.card}>
                <PanelState
                  compact
                  tone="info"
                  title="Gestor de etiquetas desativado"
                  description="Ativa a funcionalidade em Settings > Grupos para gerir o catalogo central."
                />
                <div style={S.inlineRow}>
                  <button type="button" style={S.primaryBtn} onClick={() => openSettingsSection("groups")}>
                    <Icons.Settings size={12} />
                    Abrir Settings
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
          </div>
        </div>
      ) : null}

      <div style={S.viewport}>
        <div style={{ ...S.track, transform: trackTransform }}>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <div>
                <div style={S.panelTitle}>Grupos</div>
                <div style={S.panelHint}>Criar, filtrar, arquivar e selecionar grupos.</div>
              </div>
            </div>

            <div style={S.card}>
              <div style={S.fieldLabel}>Novo grupo</div>
              <div style={S.inlineRow}>
                <input style={S.input} value={newGroupName} onChange={(event) => setNewGroupName(event.target.value)} placeholder="Nome do grupo" />
                <button type="button" style={S.primaryBtn} onClick={() => void handleCreateGroup()} disabled={busy}>
                  <Icons.Plus size={12} />
                  Criar
                </button>
              </div>
              <input
                style={S.input}
                value={newGroupLabels}
                onChange={(event) => setNewGroupLabels(event.target.value)}
                placeholder="Etiquetas iniciais (ganho, marca, cliente...)"
              />
              {labelsManagerEnabled && allLabels.length ? (
                <div style={S.tagChooser}>
                  {allLabels.map((label) => {
                    const active = parseLabels(newGroupLabels).some((entry) => normalizeText(entry) === normalizeText(label));
                    return (
                      <button
                        key={`new:${label}`}
                        type="button"
                        style={active ? S.chipActive : S.chip}
                        onClick={() => setNewGroupLabels((current) => toggleLabelsTextValue(current, label))}
                      >
                        {label}
                      </button>
                    );
                  })}
                </div>
              ) : null}
            </div>

            <div style={S.card}>
              <div style={S.fieldLabel}>Filtros</div>
              <input style={S.input} value={groupQuery} onChange={(event) => setGroupQuery(event.target.value)} placeholder="Pesquisar grupos" />
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
                return (
                  <button key={group.id} type="button" style={selected ? S.groupItemActive : S.groupItem} onClick={() => { setSelectedGroupId(group.id); setView("detail"); }}>
                    <div style={S.groupItemHead}>
                      <div style={S.groupName}>{group.name}</div>
                      <span style={{ ...S.statusBadge, ...(group.status === "concluido" ? S.statusDone : group.status === "em_progresso" ? S.statusProgress : S.statusAnalysis) }}>
                        {statusLabel(group.status)}
                      </span>
                    </div>
                    {group.description ? <div style={S.groupDescription}>{group.description}</div> : null}
                    <div style={S.groupMeta}>
                      <span>{group.memberCount || 0} email(s)</span>
                      <span>{group.documentCount || 0} doc(s)</span>
                      {group.isArchived ? <span>Arquivado</span> : null}
                    </div>
                    {group.labels?.length ? (
                      <div style={S.groupLabels}>
                        {group.labels.slice(0, 4).map((label) => (
                          <span key={label} style={S.labelBadge}>{label}</span>
                        ))}
                      </div>
                    ) : null}
                  </button>
                );
              })}
            </div>
          </section>
          <section style={S.panel}>
            <div style={S.panelHeader}>
              <button type="button" style={S.backBtn} onClick={() => setView("groups")}>Voltar</button>
              <div>
                <div style={S.panelTitle}>{selectedGroup ? selectedGroup.name : "Detalhe do grupo"}</div>
                <div style={S.panelHint}>Estado, etiquetas e emails associados.</div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Seleciona um grupo" description="Escolhe um grupo na lista para o gerir." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.fieldLabel}>Dados do grupo</div>
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
                  <input style={S.input} value={draft.labelsText} onChange={(event) => setDraft((current) => ({ ...current, labelsText: event.target.value }))} placeholder="Etiquetas separadas por virgula" />
                  {labelsManagerEnabled && allLabels.length ? (
                    <div style={S.tagChooser}>
                      {allLabels.map((label) => {
                        const active = parseLabels(draft.labelsText).some((entry) => normalizeText(entry) === normalizeText(label));
                        return (
                          <button
                            key={`draft:${label}`}
                            type="button"
                            style={active ? S.chipActive : S.chip}
                            onClick={() => setDraft((current) => ({ ...current, labelsText: toggleLabelsTextValue(current.labelsText, label) }))}
                          >
                            {label}
                          </button>
                        );
                      })}
                    </div>
                  ) : null}
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
                      <div style={S.fieldLabel}>Emails do grupo</div>
                      <div style={S.smallMeta}>
                        {groupEmails.length} email(s) ligados
                        {currentEmailAlreadyLinked ? " · email atual ligado" : ""}
                      </div>
                    </div>
                    <div style={S.inlineActions}>
                      <select
                        style={S.compactSelect}
                        value={linkKind}
                        onChange={(event) => setLinkKind(event.target.value as MembershipKind)}
                      >
                        {MEMBERSHIP_OPTIONS.map((option) => (
                          <option key={option.value} value={option.value}>{option.label}</option>
                        ))}
                      </select>
                      <button type="button" style={S.secondaryBtn} onClick={() => void handleAddCurrentEmail()} disabled={busy || currentEmailAlreadyLinked}>
                        <Icons.Link size={12} />
                        Adicionar email atual
                      </button>
                      <button type="button" style={S.secondaryBtn} onClick={() => setView("library")} disabled={busy}>
                        <Icons.Plus size={12} />
                        Associar emails
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
                  {!groupEmailsLoading && !visibleGroupEmails.length ? <PanelState compact tone="info" title="Sem emails ligados" description="Usa 'Adicionar email atual' ou 'Associar emails'." /> : null}

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
                <div style={S.panelTitle}>Biblioteca de emails</div>
                <div style={S.panelHint}>Seleciona emails registados e adiciona-os ao grupo.</div>
              </div>
            </div>

            {!selectedGroup ? (
              <PanelState compact tone="info" title="Sem grupo selecionado" description="Escolhe primeiro um grupo." />
            ) : (
              <>
                <div style={S.card}>
                  <div style={S.fieldLabel}>Pesquisar emails conhecidos</div>
                  <div style={S.inlineRow}>
                    <input style={S.input} value={libraryQuery} onChange={(event) => setLibraryQuery(event.target.value)} placeholder="Assunto, remetente, grupo, registo..." />
                    <select
                      style={S.compactSelect}
                      value={linkKind}
                      onChange={(event) => setLinkKind(event.target.value as MembershipKind)}
                    >
                      {MEMBERSHIP_OPTIONS.map((option) => (
                        <option key={option.value} value={option.value}>{option.label}</option>
                      ))}
                    </select>
                    <button type="button" style={S.primaryBtn} onClick={() => void handleAddSelectedLibrary()} disabled={busy || !selectedLibraryRows.length}>
                      <Icons.Link size={12} />
                      Associar selecionados
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
      </div>
    </div>
  );
};

const baseButton: React.CSSProperties = {
  borderRadius: 12,
  border: "1px solid var(--iccc-card-border)",
  padding: "8px 12px",
  fontSize: 12,
  fontWeight: 700,
  cursor: "pointer",
  display: "inline-flex",
  alignItems: "center",
  justifyContent: "center",
  gap: 6,
};

const S: Record<string, React.CSSProperties> = {
  root: { display: "grid", gap: 12, minHeight: "100%", position: "relative" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, padding: 12, borderRadius: 18, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  headerActions: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  kicker: { fontSize: 11, fontWeight: 800, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.06em" },
  title: { fontSize: 18, fontWeight: 800, color: "var(--iccc-text)" },
  settingsShell: { display: "grid", gridTemplateColumns: "220px minmax(0, 1fr)", gap: 0, borderRadius: 20, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)", overflow: "hidden" },
  settingsMenu: { display: "grid", alignContent: "start", gap: 8, padding: 12, background: "rgba(255,255,255,0.78)", borderRight: "1px solid var(--iccc-card-border)" },
  settingsMenuItem: { ...baseButton, width: "100%", justifyContent: "flex-start", background: "transparent", color: "var(--iccc-text)" },
  settingsMenuItemActive: { ...baseButton, width: "100%", justifyContent: "flex-start", background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  settingsContent: { display: "grid", gap: 12, padding: 12, alignContent: "start" },
  settingsContentHead: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: 10 },
  settingsGrid: { display: "grid", gap: 12 },
  settingsColumns: { display: "grid", gridTemplateColumns: "minmax(220px, 260px) minmax(0, 1fr)", gap: 12, alignItems: "start" },
  managerList: { display: "grid", gap: 8, padding: 12, borderRadius: 16, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", alignContent: "start" },
  managerRow: { width: "100%", borderRadius: 12, border: "1px solid var(--iccc-card-border)", background: "#fff", padding: "10px 12px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 13, fontWeight: 700 },
  managerRowActive: { width: "100%", borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: "10px 12px", display: "flex", alignItems: "center", justifyContent: "space-between", gap: 10, cursor: "pointer", color: "var(--iccc-text)", fontSize: 13, fontWeight: 700 },
  managerCount: { display: "inline-flex", alignItems: "center", justifyContent: "center", minWidth: 26, height: 26, borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 800 },
  viewport: { overflow: "hidden", borderRadius: 20, border: "1px solid var(--iccc-card-border)", background: "var(--iccc-card-bg)", boxShadow: "var(--iccc-shadow)" },
  track: { width: "300%", display: "flex", transition: "transform 0.22s ease" },
  panel: { width: "33.3333%", padding: 12, display: "grid", alignContent: "start", gap: 12, minHeight: "calc(100vh - 220px)", boxSizing: "border-box" },
  panelHeader: { display: "flex", alignItems: "center", gap: 10 },
  panelTitle: { fontSize: 16, fontWeight: 800, color: "var(--iccc-text)" },
  panelHint: { fontSize: 12, color: "var(--iccc-text-muted)" },
  backBtn: { ...baseButton, background: "transparent", color: "var(--iccc-text)" },
  primaryBtn: { ...baseButton, background: "linear-gradient(180deg, rgba(96, 165, 250, 0.95) 0%, rgba(37, 99, 235, 0.95) 100%)", color: "#fff", border: "1px solid rgba(37, 99, 235, 0.35)" },
  secondaryBtn: { ...baseButton, background: "rgba(255,255,255,0.78)", color: "var(--iccc-text)" },
  secondaryBtnActive: { ...baseButton, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  dangerBtn: { ...baseButton, background: "rgba(254, 226, 226, 0.95)", color: "#b91c1c", border: "1px solid rgba(239, 68, 68, 0.25)" },
  card: { display: "grid", gap: 8, padding: 12, borderRadius: 16, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)" },
  fieldLabel: { fontSize: 11, fontWeight: 800, textTransform: "uppercase", color: "var(--iccc-text-muted)", letterSpacing: "0.05em" },
  input: { width: "100%", borderRadius: 12, border: "1px solid var(--iccc-card-border)", padding: "10px 12px", background: "#fff", fontSize: 13, color: "var(--iccc-text)", boxSizing: "border-box" },
  compactSelect: { borderRadius: 12, border: "1px solid var(--iccc-card-border)", padding: "8px 10px", background: "#fff", fontSize: 12, color: "var(--iccc-text)", minWidth: 120 },
  textarea: { width: "100%", minHeight: 86, borderRadius: 12, border: "1px solid var(--iccc-card-border)", padding: "10px 12px", background: "#fff", fontSize: 13, color: "var(--iccc-text)", boxSizing: "border-box", resize: "vertical" },
  select: { width: "100%", borderRadius: 12, border: "1px solid var(--iccc-card-border)", padding: "10px 12px", background: "#fff", fontSize: 13, color: "var(--iccc-text)" },
  inlineRow: { display: "flex", gap: 8, alignItems: "center" },
  toggleRow: { display: "inline-flex", gap: 8, alignItems: "center", fontSize: 12, color: "var(--iccc-text)" },
  chipWrap: { display: "flex", flexWrap: "wrap", gap: 8 },
  tagChooser: { display: "flex", flexWrap: "wrap", gap: 8, marginTop: 2 },
  chip: { ...baseButton, padding: "6px 10px", fontSize: 11, background: "rgba(255,255,255,0.85)", color: "var(--iccc-text)" },
  chipActive: { ...baseButton, padding: "6px 10px", fontSize: 11, background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", border: "1px solid rgba(37, 99, 235, 0.2)" },
  listWrap: { display: "grid", gap: 8 },
  groupItem: { width: "100%", textAlign: "left", borderRadius: 16, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.72)", padding: 12, display: "grid", gap: 8, cursor: "pointer" },
  groupItemActive: { width: "100%", textAlign: "left", borderRadius: 16, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 12, display: "grid", gap: 8, cursor: "pointer" },
  groupItemHead: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 },
  groupName: { fontSize: 14, fontWeight: 800, color: "var(--iccc-text)" },
  groupDescription: { fontSize: 12, color: "var(--iccc-text-muted)", lineHeight: 1.45 },
  groupMeta: { display: "flex", gap: 10, flexWrap: "wrap", fontSize: 11, color: "var(--iccc-text-muted)" },
  groupLabels: { display: "flex", gap: 6, flexWrap: "wrap" },
  labelBadge: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 11, fontWeight: 700 },
  principalBadge: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(37, 99, 235, 0.08)", color: "#1d4ed8", fontSize: 10, fontWeight: 800, textTransform: "uppercase" },
  refBadge: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 10, fontWeight: 800, textTransform: "uppercase" },
  recordBadge: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, background: "rgba(15, 23, 42, 0.06)", color: "var(--iccc-text)", fontSize: 11, fontWeight: 700 },
  statusBadge: { display: "inline-flex", alignItems: "center", padding: "4px 8px", borderRadius: 999, border: "1px solid transparent", fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.04em" },
  statusAnalysis: { background: "rgba(217, 119, 6, 0.12)", color: "#b45309", borderColor: "rgba(217, 119, 6, 0.2)" },
  statusProgress: { background: "rgba(37, 99, 235, 0.12)", color: "#1d4ed8", borderColor: "rgba(37, 99, 235, 0.2)" },
  statusDone: { background: "rgba(16, 185, 129, 0.12)", color: "#047857", borderColor: "rgba(16, 185, 129, 0.25)" },
  sectionRow: { display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start" },
  smallMeta: { fontSize: 11, color: "var(--iccc-text-muted)" },
  inlineActions: { display: "flex", gap: 8, flexWrap: "wrap", justifyContent: "flex-end" },
  emailList: { display: "grid", gap: 8 },
  emailRow: { display: "grid", gridTemplateColumns: "24px minmax(0, 1fr) 34px", gap: 8, alignItems: "center", borderRadius: 14, border: "1px solid var(--iccc-card-border)", background: "rgba(255,255,255,0.82)", padding: 10 },
  emailRowActive: { display: "grid", gridTemplateColumns: "24px minmax(0, 1fr) 34px", gap: 8, alignItems: "center", borderRadius: 14, border: "1px solid rgba(37, 99, 235, 0.28)", background: "rgba(219, 234, 254, 0.72)", padding: 10 },
  checkboxCell: { display: "inline-flex", alignItems: "center", justifyContent: "center" },
  emailMain: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 4, cursor: "pointer", minWidth: 0 },
  emailSubject: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap", fontSize: 13, fontWeight: 800, color: "var(--iccc-text)" },
  emailMeta: { fontSize: 11, color: "var(--iccc-text-muted)" },
  emailAssociations: { display: "flex", gap: 6, flexWrap: "wrap" },
  currentEmailBadge: { display: "inline-flex", alignItems: "center", padding: "3px 7px", borderRadius: 999, background: "rgba(16, 185, 129, 0.12)", color: "#047857", fontSize: 10, fontWeight: 800, textTransform: "uppercase" },
  iconOnlyBtn: { ...baseButton, width: 34, height: 34, padding: 0, background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)" },
};

export default GroupManagerCockpit;
