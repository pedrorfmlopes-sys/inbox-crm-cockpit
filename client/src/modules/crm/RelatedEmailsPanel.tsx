import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  addEmailToLinkGroup,
  createLinkGroup,
  getGroupEmails,
  getLinksByRecord,
  getOdooAutoLoginUrl,
  getRelatedEmailContext,
  listLinkGroups,
  searchOdoo,
  type LinkEntry,
  type LinkGroupEntry,
  type OdooMeta,
  type RelatedEmailEntry,
} from "@/api";
import { openLinkedOutlookEmail, type OutlookMessageContext } from "@/office";
import { type CockpitSettingsV1 } from "@/settings";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type SupportedModel = "res.partner" | "crm.lead" | "project.task" | "project.project" | "helpdesk.ticket";
type ExploreMode = "records" | "groups";
type EntityOption = { model: SupportedModel; label: string };
type RecordOption = { id: number; name: string };

const ENTITY_OPTIONS: EntityOption[] = [
  { model: "res.partner", label: "Contacto" },
  { model: "crm.lead", label: "Lead" },
  { model: "project.task", label: "Tarefa" },
  { model: "project.project", label: "Projeto" },
  { model: "helpdesk.ticket", label: "Ticket" },
];

function normalizeMessageId(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function getEntityLabel(model: string): string {
  return ENTITY_OPTIONS.find((option) => option.model === model)?.label || model;
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

function buildRecordUrl(settings: CockpitSettingsV1 | null, meta: OdooMeta | null, model: string, recordId: number): string {
  const baseUrl = settings?.odooUrl || meta?.baseUrl || meta?.webBaseUrl || meta?.url || "";
  if (!baseUrl) return "";
  const db = settings?.odooDb || meta?.db || "";
  const targetBase = db ? `/web?db=${encodeURIComponent(db)}` : "/web";
  const target = `${targetBase}#id=${encodeURIComponent(String(recordId))}&model=${encodeURIComponent(model)}&view_type=form`;
  return getOdooAutoLoginUrl(settings?.odooSessionToken || null, target, baseUrl);
}

function dedupeRecordLinks(entries: LinkEntry[]): LinkEntry[] {
  const seen = new Set<string>();
  return (entries || []).filter((entry) => {
    const recordId = Number(entry.recordId || entry.resId || 0);
    const key = `${String(entry.model || "").trim()}:${recordId}`;
    if (!entry.model || !recordId || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function getCurrentEmailPayload(currentCtx: OutlookMessageContext) {
  return {
    itemId: String(currentCtx.itemId || "").trim(),
    internetMessageId: normalizeMessageId(currentCtx.internetMessageId),
    conversationId: String(currentCtx.conversationId || "").trim(),
    subject: String(currentCtx.subject || "").trim(),
    fromEmail: String(currentCtx.fromEmail || "").trim(),
    fromName: String(currentCtx.fromName || "").trim(),
    receivedAtIso: String(currentCtx.receivedDateTimeIso || "").trim(),
    messageDateIso: String(currentCtx.receivedDateTimeIso || "").trim(),
  };
}

function hasCurrentEmailIdentity(currentCtx: OutlookMessageContext): boolean {
  const payload = getCurrentEmailPayload(currentCtx);
  return Boolean(payload.itemId || payload.internetMessageId || payload.conversationId);
}

function CompactSection({
  title,
  subtitle,
  count,
  defaultOpen = true,
  actions,
  children,
}: {
  title: string;
  subtitle?: string;
  count?: string | number;
  defaultOpen?: boolean;
  actions?: React.ReactNode;
  children: React.ReactNode;
}) {
  const [open, setOpen] = useState(defaultOpen);
  return (
    <section style={styles.compactSection}>
      <div style={styles.compactHeader}>
        <button type="button" style={styles.compactToggle} onClick={() => setOpen((value) => !value)} title={open ? "Recolher secao" : "Expandir secao"}>
          {open ? <Icons.ArrowUp size={11} /> : <Icons.ArrowDown size={11} />}
          <span style={styles.compactHeaderText}>
            <span style={styles.compactTitle}>{title}</span>
            {subtitle ? <span style={styles.compactSubtitle}>{subtitle}</span> : null}
          </span>
        </button>
        <div style={styles.compactHeaderSide}>
          {typeof count !== "undefined" ? <span style={styles.countPill}>{count}</span> : null}
          {actions}
        </div>
      </div>
      {open ? <div style={styles.compactBody}>{children}</div> : null}
    </section>
  );
}

function IconActionButton({
  title,
  onClick,
  icon,
  disabled,
  tone = "default",
}: {
  title: string;
  onClick?: () => void;
  icon: React.ReactNode;
  disabled?: boolean;
  tone?: "default" | "primary";
}) {
  return (
    <button type="button" title={title} aria-label={title} style={tone === "primary" ? styles.iconBtnPrimary : styles.iconBtn} onClick={onClick} disabled={disabled}>
      {icon}
    </button>
  );
}

function RecordSearchPicker({
  model,
  selected,
  onSelect,
}: {
  model: SupportedModel;
  selected: RecordOption | null;
  onSelect: (record: RecordOption | null) => void;
}) {
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<RecordOption[]>([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const loadSeqRef = useRef(0);
  const effectiveText = selected ? selected.name : query;

  useEffect(() => {
    setQuery("");
    setItems([]);
    setOpen(false);
    setError(null);
  }, [model]);

  useEffect(() => {
    if (!open) return;
    const timer = window.setTimeout(async () => {
      const reqId = ++loadSeqRef.current;
      setBusy(true);
      setError(null);
      try {
        const rows = await searchOdoo(model, selected ? "" : query, 15);
        if (reqId !== loadSeqRef.current) return;
        setItems(
          (Array.isArray(rows) ? rows : [])
            .map((row: any) => ({ id: Number(row.id || 0), name: row.display_name || row.name || `#${row.id}` }))
            .filter((row) => row.id)
        );
      } catch (loadError: any) {
        if (reqId !== loadSeqRef.current) return;
        setItems([]);
        setError(loadError?.message || "Nao foi possivel carregar registos.");
      } finally {
        if (reqId === loadSeqRef.current) setBusy(false);
      }
    }, 220);
    return () => window.clearTimeout(timer);
  }, [model, open, query, selected]);

  return (
    <div style={styles.pickerShell}>
      <div style={styles.fieldLabel}>Registo Odoo</div>
      <div style={styles.pickerInputRow}>
        <input
          style={styles.pickerInput}
          value={effectiveText}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onChange={(event) => {
            const value = event.target.value;
            if (selected) onSelect(null);
            setQuery(value);
            setOpen(true);
          }}
          placeholder={`Pesquisar ${getEntityLabel(model).toLowerCase()}...`}
        />
        {selected ? (
          <IconActionButton title="Limpar selecao" onClick={() => { onSelect(null); setQuery(""); setOpen(true); }} icon={<Icons.RotateCcw size={12} />} />
        ) : (
          <IconActionButton title="Abrir pesquisa" onClick={() => setOpen(true)} icon={<Icons.ArrowRight size={12} />} tone="primary" />
        )}
      </div>
      {selected ? <div style={styles.selectedHint}>Selecionado: {selected.name} (#{selected.id})</div> : null}
      {error ? <div style={styles.errorHint}>{error}</div> : null}
      {open ? (
        <div style={styles.pickList}>
          {busy && !items.length ? <div style={styles.pickEmpty}>A procurar...</div> : null}
          {!busy && !items.length ? <div style={styles.pickEmpty}>Sem resultados.</div> : null}
          {items.map((item) => (
            <button
              key={item.id}
              style={styles.pickItem}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                onSelect(item);
                setOpen(false);
                setQuery("");
              }}
            >
              <span style={styles.pickName}>{item.name}</span>
              <span style={styles.pickMeta}>#{item.id}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function GroupSearchPicker({
  selected,
  onSelect,
}: {
  selected: LinkGroupEntry | null;
  onSelect: (group: LinkGroupEntry | null) => void;
}) {
  const [query, setQuery] = useState("");
  const [items, setItems] = useState<LinkGroupEntry[]>([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const loadSeqRef = useRef(0);
  const effectiveText = selected ? selected.name : query;

  useEffect(() => {
    if (!open) return;
    const timer = window.setTimeout(async () => {
      const reqId = ++loadSeqRef.current;
      setBusy(true);
      setError(null);
      try {
        const groups = await listLinkGroups(selected ? "" : query);
        if (reqId !== loadSeqRef.current) return;
        setItems(Array.isArray(groups) ? groups : []);
      } catch (loadError: any) {
        if (reqId !== loadSeqRef.current) return;
        setItems([]);
        setError(loadError?.message || "Nao foi possivel carregar grupos.");
      } finally {
        if (reqId === loadSeqRef.current) setBusy(false);
      }
    }, 220);
    return () => window.clearTimeout(timer);
  }, [open, query, selected]);

  return (
    <div style={styles.pickerShell}>
      <div style={styles.fieldLabel}>Grupo manual</div>
      <div style={styles.pickerInputRow}>
        <input
          style={styles.pickerInput}
          value={effectiveText}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onChange={(event) => {
            const value = event.target.value;
            if (selected) onSelect(null);
            setQuery(value);
            setOpen(true);
          }}
          placeholder="Pesquisar grupo..."
        />
        {selected ? (
          <IconActionButton title="Limpar grupo" onClick={() => { onSelect(null); setQuery(""); setOpen(true); }} icon={<Icons.RotateCcw size={12} />} />
        ) : (
          <IconActionButton title="Abrir grupos" onClick={() => setOpen(true)} icon={<Icons.ArrowRight size={12} />} tone="primary" />
        )}
      </div>
      {selected ? <div style={styles.selectedHint}>Selecionado: {selected.name}</div> : null}
      {error ? <div style={styles.errorHint}>{error}</div> : null}
      {open ? (
        <div style={styles.pickList}>
          {busy && !items.length ? <div style={styles.pickEmpty}>A procurar...</div> : null}
          {!busy && !items.length ? <div style={styles.pickEmpty}>Sem grupos.</div> : null}
          {items.map((item) => (
            <button
              key={item.id}
              style={styles.pickItem}
              onMouseDown={(event) => event.preventDefault()}
              onClick={() => {
                onSelect(item);
                setOpen(false);
                setQuery("");
              }}
            >
              <span style={styles.pickName}>{item.name}</span>
              <span style={styles.pickMeta}>{item.memberCount || 0} email(s)</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function RelatedEmailRow({
  item,
  settings,
  meta,
  onOpenEmail,
  onEditRecord,
}: {
  item: RelatedEmailEntry;
  settings: CockpitSettingsV1 | null;
  meta: OdooMeta | null;
  onOpenEmail: (item: RelatedEmailEntry) => void;
  onEditRecord: (model: string, recordId: number) => void;
}) {
  const primaryRecord = item.relatedRecords?.[0];
  const odooUrl = primaryRecord ? buildRecordUrl(settings, meta, primaryRecord.model, primaryRecord.recordId) : "";
  const hasOutlookOpen = Boolean(item.itemId || item.emailWebLink);
  return (
    <div style={styles.emailCard}>
      <div style={styles.emailHeaderRow}>
        <div style={styles.emailBodyCopy}>
          <div style={styles.emailSubject}>{item.subject || "(sem assunto)"}</div>
          <div style={styles.emailMetaRow}>
            <span>{item.fromName || item.fromEmail || "(sem remetente)"}</span>
            {formatDate(item.messageDateIso || item.receivedAtIso) ? <span>{formatDate(item.messageDateIso || item.receivedAtIso)}</span> : null}
          </div>
        </div>
        <div style={styles.rowActions}>
          <IconActionButton title={hasOutlookOpen ? "Abrir email no Outlook" : "Sem abertura direta disponivel"} onClick={hasOutlookOpen ? () => onOpenEmail(item) : undefined} icon={<Icons.MessageSquare size={12} />} tone="primary" disabled={!hasOutlookOpen} />
          {primaryRecord ? <IconActionButton title={`Editar ${getEntityLabel(primaryRecord.model).toLowerCase()} no Cockpit`} onClick={() => onEditRecord(primaryRecord.model, primaryRecord.recordId)} icon={<Icons.Edit size={12} />} /> : null}
          {odooUrl ? <a href={odooUrl} target="_blank" rel="noreferrer" style={styles.iconLinkBtn} title="Abrir no Odoo"><Icons.ExternalLink size={12} /></a> : null}
        </div>
      </div>
      {item.relatedRecords?.length ? <div style={styles.tagRow}>{item.relatedRecords.map((record) => <span key={`${record.model}:${record.recordId}`} style={styles.entityTag}>{getEntityLabel(record.model)}: {record.recordName || `#${record.recordId}`}</span>)}</div> : null}
      {item.relatedGroups?.length ? <div style={styles.tagRow}>{item.relatedGroups.map((group) => <span key={group.id} style={group.kind === "conversation" ? styles.conversationTag : styles.groupTag}>{group.kind === "conversation" ? "Conversa" : "Grupo"}: {group.name || group.id}</span>)}</div> : null}
    </div>
  );
}

export function RelatedEmailsPanel({
  currentCtx,
  currentLinks,
  meta,
  settings,
  onEditRecord,
  onStatus,
}: {
  currentCtx: OutlookMessageContext;
  currentLinks: LinkEntry[];
  meta: OdooMeta | null;
  settings: CockpitSettingsV1 | null;
  onEditRecord: (model: string, recordId: number) => void;
  onStatus: (message: string) => void;
}) {
  const [view, setView] = useState<"context" | "manual">("context");
  const [exploreMode, setExploreMode] = useState<ExploreMode>("records");
  const [manualModel, setManualModel] = useState<SupportedModel>("res.partner");
  const [selectedRecord, setSelectedRecord] = useState<RecordOption | null>(null);
  const [selectedGroup, setSelectedGroup] = useState<LinkGroupEntry | null>(null);
  const [contextItems, setContextItems] = useState<RelatedEmailEntry[]>([]);
  const [contextGroups, setContextGroups] = useState<LinkGroupEntry[]>([]);
  const [manualItems, setManualItems] = useState<RelatedEmailEntry[]>([]);
  const [groupItems, setGroupItems] = useState<RelatedEmailEntry[]>([]);
  const [contextLoading, setContextLoading] = useState(false);
  const [manualLoading, setManualLoading] = useState(false);
  const [groupLoading, setGroupLoading] = useState(false);
  const [contextError, setContextError] = useState<string | null>(null);
  const [manualError, setManualError] = useState<string | null>(null);
  const [groupError, setGroupError] = useState<string | null>(null);
  const [newGroupName, setNewGroupName] = useState("");
  const [groupActionBusy, setGroupActionBusy] = useState(false);
  const [contextReloadToken, setContextReloadToken] = useState(0);
  const contextLoadSeq = useRef(0);
  const manualLoadSeq = useRef(0);
  const groupLoadSeq = useRef(0);

  const emailPayload = useMemo(
    () => getCurrentEmailPayload(currentCtx),
    [currentCtx.itemId, currentCtx.internetMessageId, currentCtx.conversationId, currentCtx.subject, currentCtx.fromEmail, currentCtx.fromName, currentCtx.receivedDateTimeIso]
  );
  const currentLinkSignature = useMemo(
    () =>
      dedupeRecordLinks(currentLinks)
        .map((link) => `${String(link.model || "").trim()}:${Number(link.recordId || link.resId || 0)}`)
        .sort()
        .join("|"),
    [currentLinks]
  );
  const linkedRecords = useMemo(() => dedupeRecordLinks(currentLinks), [currentLinks]);
  const selectedRecordUrl = useMemo(() => {
    if (!selectedRecord) return "";
    return buildRecordUrl(settings, meta, manualModel, selectedRecord.id);
  }, [manualModel, meta, selectedRecord, settings]);

  useEffect(() => {
    setSelectedRecord(null);
    setManualItems([]);
    setManualError(null);
  }, [manualModel]);

  useEffect(() => {
    const reqId = ++contextLoadSeq.current;
    if (!hasCurrentEmailIdentity(currentCtx)) {
      setContextItems([]);
      setContextGroups([]);
      setContextError(null);
      setContextLoading(false);
      return;
    }
    setContextLoading(true);
    setContextError(null);
    getRelatedEmailContext(emailPayload)
      .then((response) => {
        if (reqId !== contextLoadSeq.current) return;
        setContextItems(response.emails || []);
        setContextGroups(response.groups || []);
      })
      .catch((error: any) => {
        if (reqId !== contextLoadSeq.current) return;
        setContextItems([]);
        setContextGroups([]);
        setContextError(error?.message || "Nao foi possivel carregar o contexto relacionado.");
      })
      .finally(() => {
        if (reqId === contextLoadSeq.current) setContextLoading(false);
      });
  }, [currentCtx.itemId, currentCtx.internetMessageId, currentCtx.conversationId, currentCtx.subject, currentCtx.fromEmail, currentCtx.receivedDateTimeIso, currentLinkSignature, contextReloadToken, emailPayload]);

  useEffect(() => {
    const reqId = ++manualLoadSeq.current;
    if (!selectedRecord?.id) {
      setManualItems([]);
      setManualError(null);
      setManualLoading(false);
      return;
    }
    setManualLoading(true);
    setManualError(null);
    getLinksByRecord(manualModel, selectedRecord.id)
      .then((rows) => {
        if (reqId !== manualLoadSeq.current) return;
        setManualItems((rows || []).map((row) => ({ ...row, relatedRecords: [{ model: manualModel, recordId: selectedRecord.id, recordName: selectedRecord.name }] })));
      })
      .catch((error: any) => {
        if (reqId !== manualLoadSeq.current) return;
        setManualItems([]);
        setManualError(error?.message || "Nao foi possivel carregar emails relacionados.");
      })
      .finally(() => {
        if (reqId === manualLoadSeq.current) setManualLoading(false);
      });
  }, [manualModel, selectedRecord]);

  useEffect(() => {
    const reqId = ++groupLoadSeq.current;
    if (!selectedGroup?.id) {
      setGroupItems([]);
      setGroupError(null);
      setGroupLoading(false);
      return;
    }
    setGroupLoading(true);
    setGroupError(null);
    getGroupEmails(selectedGroup.id)
      .then((rows) => {
        if (reqId !== groupLoadSeq.current) return;
        setGroupItems(rows || []);
      })
      .catch((error: any) => {
        if (reqId !== groupLoadSeq.current) return;
        setGroupItems([]);
        setGroupError(error?.message || "Nao foi possivel carregar os emails do grupo.");
      })
      .finally(() => {
        if (reqId === groupLoadSeq.current) setGroupLoading(false);
      });
  }, [selectedGroup]);

  async function handleOpenEmail(item: RelatedEmailEntry) {
    const opened = await openLinkedOutlookEmail({ itemId: item.itemId, emailWebLink: item.emailWebLink }).catch(() => false);
    if (!opened) onStatus("Nao foi possivel abrir este email diretamente neste ambiente.");
  }

  async function linkCurrentEmailToGroup(group: LinkGroupEntry) {
    if (!group?.id || !hasCurrentEmailIdentity(currentCtx)) return;
    setGroupActionBusy(true);
    try {
      await addEmailToLinkGroup(group.id, emailPayload);
      setSelectedGroup(group);
      setContextReloadToken((value) => value + 1);
      onStatus(`Email atual ligado ao grupo "${group.name}".`);
    } catch (error: any) {
      onStatus(error?.message || "Nao foi possivel ligar este email ao grupo.");
    } finally {
      setGroupActionBusy(false);
    }
  }

  async function createAndLinkGroup() {
    const name = String(newGroupName || "").trim();
    if (!name) return;
    setGroupActionBusy(true);
    try {
      const group = await createLinkGroup({ name });
      setSelectedGroup(group);
      setNewGroupName("");
      if (hasCurrentEmailIdentity(currentCtx)) {
        await addEmailToLinkGroup(group.id, emailPayload);
        onStatus(`Grupo "${group.name}" criado e ligado ao email atual.`);
      } else {
        onStatus(`Grupo "${group.name}" criado.`);
      }
      setContextReloadToken((value) => value + 1);
    } catch (error: any) {
      onStatus(error?.message || "Nao foi possivel criar o grupo.");
    } finally {
      setGroupActionBusy(false);
    }
  }

  function renderRelatedList(items: RelatedEmailEntry[], emptyTitle: string, emptyDescription: string) {
    if (!items.length) {
      return <PanelState tone="empty" title={emptyTitle} description={emptyDescription} compact />;
    }
    return (
      <div style={styles.listShell}>
        {items.map((item) => (
          <RelatedEmailRow
            key={`${item.id || item.itemId || item.internetMessageId || item.conversationId || item.subject}::${item.messageDateIso || item.receivedAtIso || ""}`}
            item={item}
            settings={settings}
            meta={meta}
            onOpenEmail={handleOpenEmail}
            onEditRecord={onEditRecord}
          />
        ))}
      </div>
    );
  }

  function renderContextView() {
    const customGroups = contextGroups.filter((group) => group.kind === "custom");
    return (
      <div style={styles.modeShell}>
        <div style={styles.metricStrip}>
          <div style={styles.metricMini}><span style={styles.metricMiniValue}>{contextItems.length}</span><span style={styles.metricMiniLabel}>emails relacionados</span></div>
          <div style={styles.metricMini}><span style={styles.metricMiniValue}>{linkedRecords.length}</span><span style={styles.metricMiniLabel}>registos Odoo</span></div>
          <div style={styles.metricMini}><span style={styles.metricMiniValue}>{contextGroups.length}</span><span style={styles.metricMiniLabel}>grupos ativos</span></div>
        </div>

        <CompactSection
          title="Relacionados agora"
          subtitle="Uniao de contexto Odoo, conversa Outlook e grupos manuais."
          count={contextLoading ? "..." : contextItems.length}
          actions={<IconActionButton title="Atualizar contexto" onClick={() => setContextReloadToken((value) => value + 1)} icon={<Icons.RefreshCw size={12} />} />}
        >
          {!hasCurrentEmailIdentity(currentCtx) ? (
            <PanelState tone="info" title="Sem email selecionado" description="Abre um email no Outlook para carregar o contexto automatico." compact />
          ) : contextLoading ? (
            <PanelState tone="loading" title="A carregar contexto" description="Estamos a reunir emails relacionados pelo modelo central." compact />
          ) : contextError ? (
            <PanelState tone="error" title="Contexto indisponivel" description={contextError} compact />
          ) : (
            renderRelatedList(contextItems, "Sem outros emails relacionados", "Ainda nao existem outros emails relevantes ligados ao mesmo contexto.")
          )}
        </CompactSection>

        <CompactSection title="Ligacoes Odoo atuais" subtitle="Registos Odoo ligados ao email aberto." count={linkedRecords.length} defaultOpen={false}>
          {linkedRecords.length ? (
            <div style={styles.tagRow}>
              {linkedRecords.map((record) => (
                <span key={`${record.model}:${record.recordId || record.resId}`} style={styles.entityTag}>
                  {getEntityLabel(record.model)}: {record.recordName || record.name || `#${record.recordId || record.resId}`}
                </span>
              ))}
            </div>
          ) : (
            <PanelState tone="empty" title="Sem ligacoes Odoo" description="Podes continuar a usar grupos manuais mesmo sem associar este email a um registo Odoo." compact />
          )}
        </CompactSection>

        <CompactSection title="Grupos manuais" subtitle="Liga emails entre si mesmo quando nao existe processo Odoo." count={customGroups.length}>
          <div style={styles.groupManager}>
            <GroupSearchPicker selected={selectedGroup} onSelect={setSelectedGroup} />
            <div style={styles.inlineActionRow}>
              <input style={styles.pickerInput} value={newGroupName} onChange={(event) => setNewGroupName(event.target.value)} placeholder="Novo grupo..." />
              <IconActionButton title="Criar grupo" onClick={createAndLinkGroup} icon={<Icons.Plus size={12} />} tone="primary" disabled={groupActionBusy || !String(newGroupName || "").trim()} />
              <IconActionButton title="Ligar email atual ao grupo selecionado" onClick={selectedGroup ? () => linkCurrentEmailToGroup(selectedGroup) : undefined} icon={<Icons.Link size={12} />} disabled={groupActionBusy || !selectedGroup || !hasCurrentEmailIdentity(currentCtx)} />
            </div>
            {customGroups.length ? (
              <div style={styles.tagRow}>
                {customGroups.map((group) => (
                  <button type="button" key={group.id} style={selectedGroup?.id === group.id ? styles.groupChipActive : styles.groupChip} title={`Selecionar grupo ${group.name}`} onClick={() => setSelectedGroup(group)}>
                    {group.name}
                  </button>
                ))}
              </div>
            ) : (
              <PanelState tone="empty" title="Sem grupos manuais" description="Cria um grupo para ligar este email a outros fora do Odoo." compact />
            )}
          </div>
        </CompactSection>
      </div>
    );
  }

  function renderRecordExplore() {
    return (
      <div style={styles.modeShell}>
        <div style={styles.formGrid}>
          <div style={styles.fieldBlock}>
            <div style={styles.fieldLabel}>Tipo de entidade</div>
            <select style={styles.entitySelect} value={manualModel} onChange={(event) => setManualModel(event.target.value as SupportedModel)}>
              {ENTITY_OPTIONS.map((option) => <option key={option.model} value={option.model}>{option.label}</option>)}
            </select>
          </div>
          <RecordSearchPicker model={manualModel} selected={selectedRecord} onSelect={setSelectedRecord} />
        </div>

        {selectedRecord ? (
          <div style={styles.selectedCard}>
            <div style={styles.selectedCardBody}>
              <div style={styles.selectedCardLabel}>{getEntityLabel(manualModel)}</div>
              <div style={styles.selectedCardTitle}>{selectedRecord.name}</div>
            </div>
            <div style={styles.rowActions}>
              <IconActionButton title={`Editar ${getEntityLabel(manualModel).toLowerCase()} no Cockpit`} onClick={() => onEditRecord(manualModel, selectedRecord.id)} icon={<Icons.Edit size={12} />} />
              {selectedRecordUrl ? <a href={selectedRecordUrl} target="_blank" rel="noreferrer" style={styles.iconLinkBtn} title="Abrir no Odoo"><Icons.ExternalLink size={12} /></a> : null}
            </div>
          </div>
        ) : null}

        <CompactSection title="Emails do registo" subtitle="Historico persistido para a entidade escolhida." count={manualLoading ? "..." : manualItems.length}>
          {manualLoading ? (
            <PanelState tone="loading" title="A carregar emails relacionados" description="Estamos a procurar emails ligados a este registo." compact />
          ) : manualError ? (
            <PanelState tone="error" title="Nao foi possivel carregar" description={manualError} compact />
          ) : !selectedRecord ? (
            <PanelState tone="empty" title="Escolhe um registo" description="Seleciona uma entidade e um registo para ver os emails relacionados." compact />
          ) : (
            renderRelatedList(manualItems, "Sem emails relacionados", "Nao existem emails ligados a este registo no armazenamento atual.")
          )}
        </CompactSection>
      </div>
    );
  }

  function renderGroupExplore() {
    return (
      <div style={styles.modeShell}>
        <div style={styles.formGrid}>
          <GroupSearchPicker selected={selectedGroup} onSelect={setSelectedGroup} />
          <div style={styles.fieldBlock}>
            <div style={styles.fieldLabel}>Novo grupo</div>
            <div style={styles.inlineActionRow}>
              <input style={styles.pickerInput} value={newGroupName} onChange={(event) => setNewGroupName(event.target.value)} placeholder="Nome do grupo..." />
              <IconActionButton title="Criar grupo" onClick={createAndLinkGroup} icon={<Icons.Plus size={12} />} tone="primary" disabled={groupActionBusy || !String(newGroupName || "").trim()} />
            </div>
          </div>
        </div>

        {selectedGroup ? (
          <div style={styles.selectedCard}>
            <div style={styles.selectedCardBody}>
              <div style={styles.selectedCardLabel}>Grupo manual</div>
              <div style={styles.selectedCardTitle}>{selectedGroup.name}</div>
              <div style={styles.selectedHint}>{selectedGroup.memberCount || 0} email(s) registado(s)</div>
            </div>
            <div style={styles.rowActions}>
              <IconActionButton title="Recarregar grupo" onClick={() => setSelectedGroup({ ...selectedGroup })} icon={<Icons.RefreshCw size={12} />} />
              <IconActionButton title="Ligar email atual a este grupo" onClick={() => linkCurrentEmailToGroup(selectedGroup)} icon={<Icons.Link size={12} />} tone="primary" disabled={groupActionBusy || !hasCurrentEmailIdentity(currentCtx)} />
            </div>
          </div>
        ) : null}

        <CompactSection title="Emails do grupo" subtitle="Vista manual para bundles de emails sem Odoo." count={groupLoading ? "..." : groupItems.length}>
          {groupLoading ? (
            <PanelState tone="loading" title="A carregar grupo" description="Estamos a reunir os emails deste grupo." compact />
          ) : groupError ? (
            <PanelState tone="error" title="Grupo indisponivel" description={groupError} compact />
          ) : !selectedGroup ? (
            <PanelState tone="empty" title="Escolhe um grupo" description="Seleciona um grupo manual para ver os emails ligados." compact />
          ) : (
            renderRelatedList(groupItems, "Sem emails no grupo", "Este grupo ainda nao tem outros emails ligados.")
          )}
        </CompactSection>
      </div>
    );
  }

  return (
    <div style={styles.section}>
      <div style={styles.header}>
        <div style={styles.headerLead}>
          <span style={styles.headerTitle}>Emails relacionados</span>
          <span style={styles.headerHint}>Modelo central de contexto, conversa e grupos.</span>
        </div>
        <div style={styles.headerActions}>
          <button style={view === "context" ? styles.switchBtnActive : styles.switchBtn} onClick={() => setView("context")} title="Ver contexto do email atual"><Icons.Activity size={12} />Contexto</button>
          <button style={view === "manual" ? styles.switchBtnActive : styles.switchBtn} onClick={() => setView("manual")} title="Explorar manualmente"><Icons.Database size={12} />Explorar</button>
        </div>
      </div>

      <div style={styles.content}>
        {view === "context" ? renderContextView() : (
          <div style={styles.modeShell}>
            <div style={styles.exploreSwitch}>
              <button style={exploreMode === "records" ? styles.exploreBtnActive : styles.exploreBtn} onClick={() => setExploreMode("records")} title="Explorar por entidades Odoo"><Icons.Database size={12} />Odoo</button>
              <button style={exploreMode === "groups" ? styles.exploreBtnActive : styles.exploreBtn} onClick={() => setExploreMode("groups")} title="Explorar por grupos manuais"><Icons.Link size={12} />Grupos</button>
            </div>
            {exploreMode === "records" ? renderRecordExplore() : renderGroupExplore()}
          </div>
        )}
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  section: { border: "1px solid #DFE1E6", borderRadius: "6px", overflow: "hidden", background: "#FFFFFF" },
  header: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "10px", flexWrap: "wrap", padding: "8px 10px", background: "#F7F8FA", borderBottom: "1px solid #DFE1E6" },
  headerLead: { display: "grid", gap: "2px", minWidth: 0 },
  headerTitle: { fontSize: "11px", fontWeight: 800, color: "#172B4D", textTransform: "uppercase", letterSpacing: "0.05em" },
  headerHint: { fontSize: "11px", color: "#6B778C" },
  headerActions: { display: "flex", gap: "6px", flexWrap: "wrap" },
  switchBtn: { border: "1px solid #C1C7D0", background: "#FFFFFF", color: "#42526E", borderRadius: "14px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "5px" },
  switchBtnActive: { border: "1px solid #0052CC", background: "#DEEBFF", color: "#0747A6", borderRadius: "14px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "5px" },
  content: { padding: "10px", minWidth: 0 },
  modeShell: { display: "grid", gap: "10px", minWidth: 0 },
  exploreSwitch: { display: "inline-flex", gap: "6px", flexWrap: "wrap" },
  exploreBtn: { border: "1px solid #DFE1E6", background: "#FFFFFF", color: "#42526E", borderRadius: "999px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "5px" },
  exploreBtnActive: { border: "1px solid #0C66E4", background: "#E9F2FF", color: "#0C66E4", borderRadius: "999px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "5px" },
  metricStrip: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: "8px" },
  metricMini: { border: "1px solid #DFE1E6", borderRadius: "6px", background: "#FAFBFC", padding: "8px", display: "grid", gap: "2px" },
  metricMiniValue: { fontSize: "16px", fontWeight: 800, color: "#172B4D" },
  metricMiniLabel: { fontSize: "10px", color: "#6B778C", textTransform: "uppercase", letterSpacing: "0.04em" },
  compactSection: { border: "1px solid #DFE1E6", borderRadius: "6px", overflow: "hidden", background: "#FFFFFF" },
  compactHeader: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: "8px", flexWrap: "wrap", padding: "7px 9px", background: "#FAFBFC", borderBottom: "1px solid #EBECF0" },
  compactToggle: { border: "none", background: "transparent", padding: 0, cursor: "pointer", display: "inline-flex", alignItems: "center", gap: "6px", minWidth: 0, textAlign: "left" },
  compactHeaderText: { display: "grid", gap: "2px", minWidth: 0 },
  compactTitle: { fontSize: "11px", fontWeight: 800, color: "#172B4D", textTransform: "uppercase", letterSpacing: "0.05em" },
  compactSubtitle: { fontSize: "11px", color: "#6B778C", lineHeight: 1.4 },
  compactHeaderSide: { display: "inline-flex", alignItems: "center", gap: "6px", flexWrap: "wrap" },
  countPill: { fontSize: "10px", fontWeight: 800, color: "#0747A6", background: "#DEEBFF", borderRadius: "999px", padding: "2px 7px" },
  compactBody: { padding: "10px", display: "grid", gap: "10px", minWidth: 0 },
  iconBtn: { border: "1px solid #DFE1E6", background: "#FFFFFF", color: "#42526E", borderRadius: "6px", width: "28px", height: "28px", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 },
  iconBtnPrimary: { border: "1px solid #0C66E4", background: "#0C66E4", color: "#FFFFFF", borderRadius: "6px", width: "28px", height: "28px", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer", flexShrink: 0 },
  iconLinkBtn: { border: "1px solid #DFE1E6", background: "#FFFFFF", color: "#0C66E4", borderRadius: "6px", width: "28px", height: "28px", display: "inline-flex", alignItems: "center", justifyContent: "center", textDecoration: "none", flexShrink: 0 },
  formGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 220px), 1fr))", gap: "10px", alignItems: "start" },
  fieldBlock: { display: "grid", gap: "4px", minWidth: 0 },
  fieldLabel: { fontSize: "10px", fontWeight: 800, color: "#42526E", textTransform: "uppercase", letterSpacing: "0.05em" },
  pickerShell: { position: "relative", display: "grid", gap: "4px", minWidth: 0 },
  pickerInputRow: { display: "flex", gap: "6px", alignItems: "center", minWidth: 0 },
  inlineActionRow: { display: "flex", gap: "6px", alignItems: "center", flexWrap: "wrap" },
  pickerInput: { flex: 1, minWidth: 0, height: "30px", border: "1px solid #C1C7D0", borderRadius: "6px", padding: "0 9px", fontSize: "12px", color: "#172B4D", background: "#FFFFFF" },
  entitySelect: { height: "30px", border: "1px solid #C1C7D0", borderRadius: "6px", padding: "0 9px", fontSize: "12px", color: "#172B4D", background: "#FFFFFF" },
  pickList: { position: "absolute", top: "60px", left: 0, right: 0, zIndex: 5, background: "#FFFFFF", border: "1px solid #C1C7D0", borderRadius: "6px", boxShadow: "0 8px 16px rgba(9, 30, 66, 0.15)", overflow: "hidden", maxHeight: "220px", overflowY: "auto" },
  pickItem: { width: "100%", border: "none", background: "#FFFFFF", padding: "7px 9px", display: "flex", justifyContent: "space-between", alignItems: "center", cursor: "pointer", textAlign: "left", borderBottom: "1px solid #F4F5F7", gap: "8px" },
  pickName: { fontSize: "12px", color: "#172B4D", fontWeight: 600, minWidth: 0, wordBreak: "break-word" },
  pickMeta: { fontSize: "10px", color: "#6B778C", flexShrink: 0 },
  pickEmpty: { padding: "9px", fontSize: "12px", color: "#6B778C" },
  selectedHint: { fontSize: "11px", color: "#6B778C" },
  errorHint: { fontSize: "11px", color: "#BF2600" },
  selectedCard: { border: "1px solid #DFE1E6", borderRadius: "6px", padding: "9px", display: "flex", justifyContent: "space-between", alignItems: "flex-start", gap: "10px", background: "#FAFBFC", flexWrap: "wrap" },
  selectedCardBody: { display: "grid", gap: "3px", minWidth: 0, flex: "1 1 180px" },
  selectedCardLabel: { fontSize: "10px", fontWeight: 800, color: "#6B778C", textTransform: "uppercase", letterSpacing: "0.05em" },
  selectedCardTitle: { fontSize: "13px", fontWeight: 700, color: "#172B4D", lineHeight: 1.35, wordBreak: "break-word" },
  listShell: { display: "grid", gap: "8px", maxHeight: "320px", overflowY: "auto", paddingRight: "2px" },
  emailCard: { border: "1px solid #DFE1E6", borderRadius: "6px", padding: "9px", display: "grid", gap: "7px", background: "#FFFFFF" },
  emailHeaderRow: { display: "flex", justifyContent: "space-between", gap: "8px", alignItems: "flex-start" },
  emailBodyCopy: { display: "grid", gap: "4px", minWidth: 0, flex: "1 1 auto" },
  emailSubject: { fontSize: "12px", fontWeight: 700, color: "#172B4D", lineHeight: 1.4, wordBreak: "break-word" },
  emailMetaRow: { display: "flex", justifyContent: "space-between", gap: "8px", flexWrap: "wrap", color: "#6B778C", fontSize: "10px" },
  rowActions: { display: "inline-flex", gap: "6px", flexWrap: "wrap", flexShrink: 0 },
  tagRow: { display: "flex", flexWrap: "wrap", gap: "6px" },
  entityTag: { fontSize: "10px", fontWeight: 800, color: "#0747A6", background: "#DEEBFF", borderRadius: "999px", padding: "2px 8px" },
  groupTag: { fontSize: "10px", fontWeight: 700, color: "#5E4DB2", background: "#ECEBFF", borderRadius: "999px", padding: "2px 8px" },
  conversationTag: { fontSize: "10px", fontWeight: 700, color: "#0052CC", background: "#E9F2FF", borderRadius: "999px", padding: "2px 8px" },
  groupManager: { display: "grid", gap: "10px" },
  groupChip: { border: "1px solid #DFE1E6", background: "#FFFFFF", color: "#42526E", borderRadius: "999px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer" },
  groupChipActive: { border: "1px solid #0C66E4", background: "#E9F2FF", color: "#0C66E4", borderRadius: "999px", padding: "4px 9px", fontSize: "11px", fontWeight: 700, cursor: "pointer" },
};
