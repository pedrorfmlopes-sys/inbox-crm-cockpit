import React, { useEffect, useMemo, useRef, useState } from "react";
import { getLinksByRecord, getOdooAutoLoginUrl, searchOdoo, type LinkEntry, type OdooMeta } from "@/api";
import { openLinkedOutlookEmail, type OutlookMessageContext } from "@/office";
import { type CockpitSettingsV1 } from "@/settings";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type SupportedModel = "res.partner" | "crm.lead" | "project.task" | "project.project" | "helpdesk.ticket";

type EntityOption = {
  model: SupportedModel;
  label: string;
};

type RecordOption = {
  id: number;
  name: string;
};

type RelatedEmailItem = LinkEntry & {
  relatedRecords: Array<{ model: string; recordId: number; recordName: string }>;
};

const ENTITY_OPTIONS: EntityOption[] = [
  { model: "res.partner", label: "Contacto" },
  { model: "crm.lead", label: "Lead" },
  { model: "project.task", label: "Tarefa" },
  { model: "project.project", label: "Projeto" },
  { model: "helpdesk.ticket", label: "Ticket" },
];

function normalizeMessageId(value: string | undefined): string {
  return String(value || "")
    .trim()
    .toLowerCase()
    .replace(/[<>\s]/g, "");
}

function getEntityLabel(model: string): string {
  return ENTITY_OPTIONS.find((option) => option.model === model)?.label || model;
}

function getEmailIdentity(link: LinkEntry): string {
  const directKey = String(link.itemId || "").trim()
    || normalizeMessageId(link.internetMessageId)
    || String(link.conversationId || "").trim();
  if (directKey) return directKey;

  const fallbackKey = [
    String(link.subject || "").trim().toLowerCase(),
    String(link.fromEmail || "").trim().toLowerCase(),
    String(link.receivedAtIso || link.linkedAt || "").trim(),
  ].join("|");

  return fallbackKey === "||" ? "" : fallbackKey;
}

function getEventDate(link: LinkEntry): string {
  return String(link.receivedAtIso || link.linkedAt || "").trim();
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

function buildRecordUrl(
  settings: CockpitSettingsV1 | null,
  meta: OdooMeta | null,
  model: string,
  recordId: number
): string {
  const baseUrl = settings?.odooUrl || meta?.baseUrl || meta?.webBaseUrl || meta?.url || "";
  if (!baseUrl) return "";
  const db = settings?.odooDb || meta?.db || "";
  const targetBase = db ? `/web?db=${encodeURIComponent(db)}` : "/web";
  const target = `${targetBase}#id=${encodeURIComponent(String(recordId))}&model=${encodeURIComponent(model)}&view_type=form`;
  return getOdooAutoLoginUrl(settings?.odooSessionToken || null, target, baseUrl);
}

function isCurrentEmail(link: LinkEntry, currentCtx: OutlookMessageContext): boolean {
  const currentItemId = String(currentCtx.itemId || "").trim();
  const currentConversationId = String(currentCtx.conversationId || "").trim();
  const currentMessageId = normalizeMessageId(currentCtx.internetMessageId);
  const linkItemId = String(link.itemId || "").trim();
  const linkMessageId = normalizeMessageId(link.internetMessageId);
  const hasPreciseCurrentIdentity = Boolean(currentItemId || currentMessageId);

  if (currentItemId && linkItemId && currentItemId === linkItemId) return true;
  if (currentMessageId && linkMessageId && currentMessageId === linkMessageId) return true;
  if (hasPreciseCurrentIdentity) return false;
  if (currentConversationId && currentConversationId === String(link.conversationId || "").trim()) return true;
  return false;
}

function aggregateContextEmails(recordLinks: LinkEntry[], currentCtx: OutlookMessageContext): RelatedEmailItem[] {
  const byEmail = new Map<string, RelatedEmailItem>();

  for (const link of recordLinks || []) {
    if (isCurrentEmail(link, currentCtx)) continue;
    const emailKey = getEmailIdentity(link);
    if (!emailKey) continue;

    const relatedRecord = {
      model: String(link.model || "").trim(),
      recordId: Number(link.recordId || link.resId || 0),
      recordName: String(link.recordName || link.name || link.title || "").trim(),
    };
    if (!relatedRecord.model || !relatedRecord.recordId) continue;

    const current = byEmail.get(emailKey);
    if (!current) {
      byEmail.set(emailKey, {
        ...link,
        relatedRecords: [relatedRecord],
      });
      continue;
    }

    const hasRecord = current.relatedRecords.some(
      (entry) => entry.model === relatedRecord.model && entry.recordId === relatedRecord.recordId
    );
    const nextRecords = hasRecord ? current.relatedRecords : [...current.relatedRecords, relatedRecord];
    const useIncoming = getEventDate(link) > getEventDate(current);

    byEmail.set(emailKey, {
      ...(useIncoming ? { ...current, ...link } : current),
      relatedRecords: nextRecords,
    });
  }

  return Array.from(byEmail.values()).sort((a, b) => getEventDate(b).localeCompare(getEventDate(a)));
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
          (Array.isArray(rows) ? rows : []).map((row: any) => ({
            id: Number(row.id || 0),
            name: row.display_name || row.name || `#${row.id}`,
          })).filter((row) => row.id)
        );
      } catch (error: any) {
        if (reqId !== loadSeqRef.current) return;
        setItems([]);
        setError(error?.message || "Nao foi possivel carregar registos.");
      } finally {
        if (reqId === loadSeqRef.current) setBusy(false);
      }
    }, 250);

    return () => window.clearTimeout(timer);
  }, [model, open, query, selected]);

  return (
    <div style={styles.pickerShell}>
      <div style={styles.pickerLabel}>Registo</div>
      <div style={styles.pickerInputRow}>
        <input
          style={styles.pickerInput}
          value={effectiveText}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onChange={(event) => {
            const value = event.target.value;
            if (selected) {
              onSelect(null);
            }
            setQuery(value);
            setOpen(true);
          }}
          placeholder={`Pesquisar ${getEntityLabel(model).toLowerCase()}...`}
        />
        {selected ? (
          <button
            style={styles.inlineGhostBtn}
            onClick={() => {
              onSelect(null);
              setQuery("");
              setOpen(true);
            }}
            title="Limpar selecao"
          >
            Limpar
          </button>
        ) : (
          <button style={styles.inlinePrimaryBtn} onClick={() => setOpen(true)} title="Pesquisar">
            Abrir
          </button>
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

function RelatedEmailRow({
  item,
  settings,
  meta,
  onOpenEmail,
}: {
  item: RelatedEmailItem;
  settings: CockpitSettingsV1 | null;
  meta: OdooMeta | null;
  onOpenEmail: (item: RelatedEmailItem) => void;
}) {
  const primaryRecord = item.relatedRecords[0];
  const odooUrl = primaryRecord
    ? buildRecordUrl(settings, meta, primaryRecord.model, primaryRecord.recordId)
    : "";

  return (
    <div style={styles.emailCard}>
      <div style={styles.emailCardTop}>
        <div style={styles.emailCardTitle}>{item.subject || "(sem assunto)"}</div>
        <button style={styles.inlinePrimaryBtn} onClick={() => onOpenEmail(item)}>
          Outlook
        </button>
      </div>

      <div style={styles.emailMetaRow}>
        <span>{item.fromName || item.fromEmail || "(sem remetente)"}</span>
        {formatDate(getEventDate(item)) ? <span>{formatDate(getEventDate(item))}</span> : null}
      </div>

      <div style={styles.emailTagRow}>
        {item.relatedRecords.map((record) => (
          <span key={`${record.model}-${record.recordId}`} style={styles.emailTag}>
            {getEntityLabel(record.model)}: {record.recordName || `#${record.recordId}`}
          </span>
        ))}
      </div>

      <div style={styles.emailActions}>
        {odooUrl ? (
          <a href={odooUrl} target="_blank" rel="noreferrer" style={styles.inlineLinkBtn}>
            <Icons.ExternalLink size={10} />
            Odoo
          </a>
        ) : null}
      </div>
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
  const [manualModel, setManualModel] = useState<SupportedModel>("res.partner");
  const [selectedRecord, setSelectedRecord] = useState<RecordOption | null>(null);
  const [contextItems, setContextItems] = useState<RelatedEmailItem[]>([]);
  const [manualItems, setManualItems] = useState<RelatedEmailItem[]>([]);
  const [contextLoading, setContextLoading] = useState(false);
  const [manualLoading, setManualLoading] = useState(false);
  const [contextError, setContextError] = useState<string | null>(null);
  const [manualError, setManualError] = useState<string | null>(null);
  const contextLoadSeq = useRef(0);
  const manualLoadSeq = useRef(0);

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
    const linkedRecords = (currentLinks || []).filter((link) => link.model && Number(link.recordId || link.resId || 0));
    const reqId = ++contextLoadSeq.current;

    if (!linkedRecords.length) {
      setContextItems([]);
      setContextError(null);
      setContextLoading(false);
      return;
    }

    setContextLoading(true);
    setContextError(null);

    Promise.all(
      linkedRecords.map(async (link) => {
        try {
          return await getLinksByRecord(String(link.model || ""), Number(link.recordId || link.resId || 0));
        } catch (error: any) {
          return [{ ...link, relatedRecords: [], loadError: error?.message || "Falha ao carregar emails." }] as any[];
        }
      })
    ).then((responses) => {
      if (reqId !== contextLoadSeq.current) return;
      const flat = responses.flat().filter((link: any) => !link?.loadError) as LinkEntry[];
      const aggregated = aggregateContextEmails(flat, currentCtx);
      setContextItems(aggregated);
      if (!aggregated.length && responses.some((rows) => rows.some((entry: any) => entry?.loadError))) {
        const firstError = responses.flat().find((entry: any) => entry?.loadError)?.loadError;
        setContextError(String(firstError || "Nao foi possivel carregar emails relacionados."));
      }
    }).catch((error: any) => {
      if (reqId !== contextLoadSeq.current) return;
      setContextItems([]);
      setContextError(error?.message || "Nao foi possivel carregar emails relacionados.");
    }).finally(() => {
      if (reqId === contextLoadSeq.current) setContextLoading(false);
    });
  }, [currentCtx, currentLinks]);

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
        setManualItems(
          (rows || []).map((row) => ({
            ...row,
            relatedRecords: [{
              model: manualModel,
              recordId: selectedRecord.id,
              recordName: selectedRecord.name,
            }],
          }))
        );
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

  async function handleOpenEmail(item: RelatedEmailItem) {
    const opened = await openLinkedOutlookEmail({
      itemId: item.itemId,
      emailWebLink: item.emailWebLink,
    }).catch(() => false);

    if (!opened) {
      onStatus("Nao foi possivel abrir este email diretamente neste ambiente.");
    }
  }

  function renderContextView() {
    if (contextLoading) {
      return (
        <PanelState
          tone="loading"
          title="A carregar emails relacionados"
          description="Estamos a reunir outros emails ligados aos mesmos registos."
          compact
        />
      );
    }

    if (!currentLinks.length) {
      return (
        <PanelState
          tone="empty"
          title="Sem registos ligados"
          description="Liga este email a um contacto, lead, tarefa, projeto ou ticket para ver o historico relacionado."
          compact
        />
      );
    }

    if (contextError && !contextItems.length) {
      return (
        <PanelState
          tone="error"
          title="Emails relacionados indisponiveis"
          description={contextError}
          compact
        />
      );
    }

    if (!contextItems.length) {
      return (
        <PanelState
          tone="info"
          title="Sem outros emails ligados"
          description="Este email ja esta ligado, mas ainda nao encontrámos outras conversas para os mesmos registos."
          compact
        />
      );
    }

    return (
      <div style={styles.listShell}>
        {contextItems.map((item) => (
          <RelatedEmailRow
            key={`${getEmailIdentity(item)}::context`}
            item={item}
            settings={settings}
            meta={meta}
            onOpenEmail={handleOpenEmail}
          />
        ))}
      </div>
    );
  }

  function renderManualView() {
    return (
      <div style={styles.manualShell}>
        <div style={styles.formRow}>
          <div style={styles.pickerLabel}>Tipo de entidade</div>
          <select
            style={styles.entitySelect}
            value={manualModel}
            onChange={(event) => setManualModel(event.target.value as SupportedModel)}
          >
            {ENTITY_OPTIONS.map((option) => (
              <option key={option.model} value={option.model}>
                {option.label}
              </option>
            ))}
          </select>
        </div>

        <RecordSearchPicker
          model={manualModel}
          selected={selectedRecord}
          onSelect={setSelectedRecord}
        />

        {selectedRecord ? (
          <div style={styles.selectedRecordCard}>
            <div>
              <div style={styles.selectedRecordLabel}>{getEntityLabel(manualModel)}</div>
              <div style={styles.selectedRecordName}>{selectedRecord.name}</div>
            </div>
            <div style={styles.selectedRecordActions}>
              <button style={styles.inlineGhostBtn} onClick={() => onEditRecord(manualModel, selectedRecord.id)}>
                <Icons.Edit size={10} />
                Editar
              </button>
              {selectedRecordUrl ? (
                <a href={selectedRecordUrl} target="_blank" rel="noreferrer" style={styles.inlineLinkBtn}>
                  <Icons.ExternalLink size={10} />
                  Abrir no Odoo
                </a>
              ) : null}
            </div>
          </div>
        ) : null}

        <div style={styles.subSectionTitle}>Emails relacionados</div>

        {manualLoading ? (
          <PanelState
            tone="loading"
            title="A carregar emails relacionados"
            description="Estamos a procurar emails ligados a este registo."
            compact
          />
        ) : manualError ? (
          <PanelState
            tone="error"
            title="Nao foi possivel carregar"
            description={manualError}
            compact
          />
        ) : !selectedRecord ? (
          <PanelState
            tone="empty"
            title="Escolhe um registo"
            description="Seleciona uma entidade e um registo para ver os emails relacionados."
            compact
          />
        ) : !manualItems.length ? (
          <PanelState
            tone="empty"
            title="Sem emails relacionados"
            description="Nao existem emails ligados a este registo no armazenamento atual."
            compact
          />
        ) : (
          <div style={styles.listShell}>
            {manualItems.map((item) => (
              <RelatedEmailRow
                key={`${getEmailIdentity(item)}::manual`}
                item={item}
                settings={settings}
                meta={meta}
                onOpenEmail={handleOpenEmail}
              />
            ))}
          </div>
        )}
      </div>
    );
  }

  return (
    <div style={styles.section}>
      <div style={styles.header}>
        <div style={styles.headerTitle}>Emails relacionados</div>
        <div style={styles.viewSwitch}>
          <button
            style={view === "context" ? styles.switchBtnActive : styles.switchBtn}
            onClick={() => setView("context")}
          >
            Contexto
          </button>
          <button
            style={view === "manual" ? styles.switchBtnActive : styles.switchBtn}
            onClick={() => setView("manual")}
          >
            Explorar
          </button>
        </div>
      </div>

      <div style={styles.content}>
        {view === "context" ? renderContextView() : renderManualView()}
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  section: {
    border: "1px solid #DFE1E6",
    borderRadius: "3px",
    overflow: "hidden",
    background: "#FFFFFF",
  },
  header: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "12px",
    flexWrap: "wrap",
    padding: "8px 12px",
    background: "#F4F5F7",
    borderBottom: "1px solid #DFE1E6",
  },
  headerTitle: {
    fontSize: "11px",
    fontWeight: 700,
    color: "#42526E",
    textTransform: "uppercase",
  },
  viewSwitch: {
    display: "flex",
    gap: "6px",
    flexWrap: "wrap",
  },
  switchBtn: {
    border: "1px solid #C1C7D0",
    background: "#FFFFFF",
    color: "#42526E",
    borderRadius: "16px",
    padding: "4px 10px",
    fontSize: "11px",
    fontWeight: 700,
    cursor: "pointer",
  },
  switchBtnActive: {
    border: "1px solid #0052CC",
    background: "#DEEBFF",
    color: "#0747A6",
    borderRadius: "16px",
    padding: "4px 10px",
    fontSize: "11px",
    fontWeight: 700,
    cursor: "pointer",
  },
  content: {
    padding: "12px",
    display: "grid",
    gap: "12px",
    minWidth: 0,
  },
  manualShell: {
    display: "grid",
    gap: "12px",
  },
  formRow: {
    display: "grid",
    gap: "4px",
  },
  pickerShell: {
    position: "relative",
    display: "grid",
    gap: "4px",
  },
  pickerLabel: {
    fontSize: "11px",
    fontWeight: 700,
    color: "#42526E",
    textTransform: "uppercase",
  },
  pickerInputRow: {
    display: "flex",
    gap: "8px",
    alignItems: "center",
    flexWrap: "wrap",
  },
  pickerInput: {
    flex: 1,
    minWidth: 0,
    height: "32px",
    border: "1px solid #C1C7D0",
    borderRadius: "3px",
    padding: "0 10px",
    fontSize: "12px",
    color: "#172B4D",
  },
  entitySelect: {
    height: "32px",
    border: "1px solid #C1C7D0",
    borderRadius: "3px",
    padding: "0 10px",
    fontSize: "12px",
    color: "#172B4D",
    background: "#FFFFFF",
  },
  pickList: {
    position: "absolute",
    top: "62px",
    left: 0,
    right: 0,
    zIndex: 3,
    background: "#FFFFFF",
    border: "1px solid #C1C7D0",
    borderRadius: "3px",
    boxShadow: "0 8px 16px rgba(9, 30, 66, 0.15)",
    overflow: "hidden",
    maxHeight: "220px",
    overflowY: "auto",
  },
  pickItem: {
    width: "100%",
    border: "none",
    background: "#FFFFFF",
    padding: "8px 10px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    cursor: "pointer",
    textAlign: "left",
    borderBottom: "1px solid #F4F5F7",
  },
  pickName: {
    fontSize: "12px",
    color: "#172B4D",
    fontWeight: 600,
  },
  pickMeta: {
    fontSize: "11px",
    color: "#6B778C",
  },
  pickEmpty: {
    padding: "10px",
    fontSize: "12px",
    color: "#6B778C",
  },
  inlinePrimaryBtn: {
    border: "1px solid #0052CC",
    background: "#0052CC",
    color: "#FFFFFF",
    borderRadius: "3px",
    padding: "6px 10px",
    fontSize: "11px",
    fontWeight: 700,
    cursor: "pointer",
    flexShrink: 0,
  },
  inlineGhostBtn: {
    border: "1px solid #C1C7D0",
    background: "#FFFFFF",
    color: "#42526E",
    borderRadius: "3px",
    padding: "6px 10px",
    fontSize: "11px",
    fontWeight: 700,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    flexShrink: 0,
  },
  inlineLinkBtn: {
    border: "1px solid #DFE1E6",
    background: "#FFFFFF",
    color: "#0052CC",
    borderRadius: "3px",
    padding: "6px 10px",
    fontSize: "11px",
    fontWeight: 700,
    textDecoration: "none",
    display: "inline-flex",
    alignItems: "center",
    gap: "4px",
    flexShrink: 0,
  },
  selectedHint: {
    fontSize: "11px",
    color: "#6B778C",
  },
  errorHint: {
    fontSize: "11px",
    color: "#BF2600",
  },
  selectedRecordCard: {
    border: "1px solid #DFE1E6",
    borderRadius: "6px",
    padding: "10px",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: "12px",
    background: "#FAFBFC",
    flexWrap: "wrap",
  },
  selectedRecordLabel: {
    fontSize: "10px",
    fontWeight: 700,
    color: "#6B778C",
    textTransform: "uppercase",
    marginBottom: "4px",
  },
  selectedRecordName: {
    fontSize: "13px",
    fontWeight: 700,
    color: "#172B4D",
    lineHeight: 1.4,
    wordBreak: "break-word",
  },
  selectedRecordActions: {
    display: "flex",
    gap: "8px",
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  subSectionTitle: {
    fontSize: "11px",
    fontWeight: 700,
    color: "#42526E",
    textTransform: "uppercase",
  },
  listShell: {
    display: "grid",
    gap: "8px",
  },
  emailCard: {
    border: "1px solid #DFE1E6",
    borderRadius: "6px",
    padding: "10px",
    display: "grid",
    gap: "8px",
    background: "#FFFFFF",
  },
  emailCardTop: {
    display: "flex",
    justifyContent: "space-between",
    gap: "10px",
    alignItems: "flex-start",
    flexWrap: "wrap",
  },
  emailCardTitle: {
    fontSize: "13px",
    fontWeight: 700,
    color: "#172B4D",
    lineHeight: 1.4,
    minWidth: 0,
    flex: "1 1 220px",
    wordBreak: "break-word",
  },
  emailMetaRow: {
    display: "flex",
    justifyContent: "space-between",
    gap: "10px",
    color: "#6B778C",
    fontSize: "11px",
    flexWrap: "wrap",
  },
  emailTagRow: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
  },
  emailTag: {
    fontSize: "10px",
    fontWeight: 700,
    color: "#0747A6",
    background: "#DEEBFF",
    borderRadius: "12px",
    padding: "2px 8px",
  },
  emailActions: {
    display: "flex",
    justifyContent: "flex-end",
    gap: "8px",
  },
};
