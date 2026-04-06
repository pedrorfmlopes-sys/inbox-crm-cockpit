import React from "react";
import * as Icons from "@/ui/icons";
import { type LinkGroupEntry, type GroupTicketEntry, type GroupTicketSeriesEntry, type RelatedEmailEntry } from "@/api";
import { 
  type ClassificationFocus, 
  type ClassificationLayoutMode, 
  type ClassificationMetaDraft, 
  type LabelDraft, 
  type EmailLabelStatus,
  type TicketEditorMode
} from "../types";
import StatusLegend from "./StatusLegend";
import { PanelState } from "@/ui/PanelState";

export interface ClassificationEditorProps {
  classificationFocus: ClassificationFocus;
  classificationLayoutMode: ClassificationLayoutMode;
  classificationSuggestionExpanded: Record<string, boolean>;
  setClassificationSuggestionExpanded: React.Dispatch<React.SetStateAction<Record<string, boolean>>>;
  suggestedExistingGroups: LinkGroupEntry[];
  principalGroupId: string;
  clearPrincipalSelection: () => void;
  selectPrincipalGroup: (group: LinkGroupEntry) => void;
  principalSearch: string;
  setPrincipalSearchValue: (val: string) => void;
  principalCanCreate: boolean;
  handleCreateGroupAndLink: (kind: "principal" | "referencia", name?: string) => Promise<void>;
  exactPrincipalSearchGroup: LinkGroupEntry | null;
  principalSearchResults: LinkGroupEntry[];
  principalGroup: LinkGroupEntry | null;
  classificationMetaDraft: ClassificationMetaDraft;
  updateClassificationMeta: (patch: Partial<ClassificationMetaDraft>) => void;
  suggestedLabelSeeds: string[];
  selectedLabels: string[];
  applySuggestedLabel: (label: string) => void;
  classificationLabelInput: string;
  setClassificationLabelInput: (val: string) => void;
  handleClassificationLabelSearchAction: () => void;
  classificationLabelCanCreate: boolean;
  filteredClassificationLabels: string[];
  removeLabel: (label: string) => void;
  addLabel: (label: string) => void;
  selectedLabelSharedStatus: EmailLabelStatus | "";
  updateLabelDraft: (label: string, patch: Partial<LabelDraft>) => void;
  labelDrafts: Record<string, LabelDraft>;
  LABEL_STATUS_OPTIONS: any[];
  normalizedTicketSearch: string;
  ticketSearchResults: GroupTicketEntry[];
  availableTicketChoices: GroupTicketEntry[];
  selectedTicket: GroupTicketEntry | null;
  selectedSeriesId: string;
  ticketEditorMode: TicketEditorMode;
  setTicketEditorMode: (mode: TicketEditorMode) => void;
  ticketSearch: string;
  setTicketSearch: (val: string) => void;
  handleSearchTickets: (query?: string, options?: { silent?: boolean }) => Promise<GroupTicketEntry[]>;
  ticketSearchBusy: boolean;
  setSelectedSeriesId: (id: string) => void;
  setSelectionTouched: React.Dispatch<React.SetStateAction<{ principal: boolean; references: boolean; ticket: boolean }>>;
  ticketSeries: GroupTicketSeriesEntry[];
  createTicketTitle: string;
  setCreateTicketTitle: (val: string) => void;
  ticketStatusDraft: string;
  setTicketStatusDraft: (status: string) => void;
  TICKET_STATUS_OPTIONS: any[];
  effectiveTicketStatus: string | undefined;
  ticketStatusLabel: string;
  selectedTicketId: string;
  applySuggestedTicket: (id: string) => void;
  clearTicketSelection: () => void;
  referenceGroups: LinkGroupEntry[];
  toggleReferenceGroup: (id: string) => void;
  referenceSearch: string;
  setReferenceSearchValue: (val: string) => void;
  referenceCanCreate: boolean;
  exactReferenceSearchGroup: LinkGroupEntry | null;
  referenceSearchResults: LinkGroupEntry[];
  referenceGroupIds: string[];
  actionBusy: boolean;
}

const ClassificationEditor: React.FC<ClassificationEditorProps> = (props) => {
  const {
    classificationFocus,
    classificationLayoutMode,
    classificationSuggestionExpanded,
    setClassificationSuggestionExpanded,
    suggestedExistingGroups,
    principalGroupId,
    clearPrincipalSelection,
    selectPrincipalGroup,
    principalSearch,
    setPrincipalSearchValue,
    principalCanCreate,
    handleCreateGroupAndLink,
    exactPrincipalSearchGroup,
    principalSearchResults,
    principalGroup,
    classificationMetaDraft,
    updateClassificationMeta,
    suggestedLabelSeeds,
    selectedLabels,
    applySuggestedLabel,
    classificationLabelInput,
    setClassificationLabelInput,
    handleClassificationLabelSearchAction,
    classificationLabelCanCreate,
    filteredClassificationLabels,
    removeLabel,
    addLabel,
    selectedLabelSharedStatus,
    updateLabelDraft,
    labelDrafts,
    LABEL_STATUS_OPTIONS,
    normalizedTicketSearch,
    ticketSearchResults,
    availableTicketChoices,
    selectedTicket,
    selectedSeriesId,
    ticketEditorMode,
    setTicketEditorMode,
    ticketSearch,
    setTicketSearch,
    handleSearchTickets,
    ticketSearchBusy,
    setSelectedSeriesId,
    setSelectionTouched,
    ticketSeries,
    createTicketTitle,
    setCreateTicketTitle,
    ticketStatusDraft,
    setTicketStatusDraft,
    TICKET_STATUS_OPTIONS,
    effectiveTicketStatus,
    ticketStatusLabel,
    selectedTicketId,
    applySuggestedTicket,
    clearTicketSelection,
    referenceGroups,
    toggleReferenceGroup,
    referenceSearch,
    setReferenceSearchValue,
    referenceCanCreate,
    exactReferenceSearchGroup,
    referenceSearchResults,
    referenceGroupIds,
    actionBusy,
  } = props;

  function renderSuggestionTray(
    kind: "principal" | "labels",
    title: string,
    chips: Array<{ key: string; label: string; active?: boolean; onClick: () => void }>
  ) {
    const visible = chips.slice(0, 3);
    const hidden = chips.slice(3);
    const expanded = classificationSuggestionExpanded[kind];
    return (
      <div style={S.editorBlock}>
        <div style={S.editorBlockHeader}>
          <div style={S.editorBlockTitle}>{title}</div>
          {hidden.length ? (
            <button
              type="button"
              style={S.chevronBtn}
              onClick={() => setClassificationSuggestionExpanded((current) => ({ ...current, [kind]: !current[kind] }))}
            >
              {expanded ? "\u2303" : "\u2304"}
            </button>
          ) : null}
        </div>
        <div style={S.chipGridCompact}>
          {visible.length ? visible.map((chip) => (
            <button key={chip.key} type="button" style={chip.active ? S.miniChipOn : S.miniChip} onClick={chip.onClick}>
              {chip.label}
            </button>
          )) : <span style={S.mutedMini}>Sem sugestoes fortes nesta leitura.</span>}
        </div>
        {hidden.length && expanded ? (
          <div style={S.editorExpandableOpen}>
            <div style={S.editorExpandableScroll}>
              {hidden.map((chip) => (
                <button key={chip.key} type="button" style={chip.active ? S.miniChipOn : S.miniChip} onClick={chip.onClick}>
                  {chip.label}
                </button>
              ))}
            </div>
          </div>
        ) : null}
      </div>
    );
  }

  function renderPrincipalEditor() {
    const suggestionChips = suggestedExistingGroups.map((group) => ({
      key: group.id,
      label: group.name || group.id,
      active: group.id === principalGroupId,
      onClick: () => {
        if (group.id === principalGroupId) clearPrincipalSelection();
        else selectPrincipalGroup(group);
      },
    }));
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Grupo principal</div>
        <div style={S.editorLead}>Escolhe ou ajusta o dossier principal do email.</div>
        {renderSuggestionTray("principal", "Sugestoes", suggestionChips)}
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Pesquisar ou criar</div>
          <div style={S.searchInlineRow}>
            <input
              style={S.input}
              value={principalSearch}
              onChange={(event) => setPrincipalSearchValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  if (principalCanCreate) {
                    void handleCreateGroupAndLink("principal", principalSearch);
                  } else if (exactPrincipalSearchGroup) {
                    selectPrincipalGroup(exactPrincipalSearchGroup);
                  }
                }
              }}
              placeholder="Escreve o nome do grupo..."
            />
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => {
                if (exactPrincipalSearchGroup) {
                  selectPrincipalGroup(exactPrincipalSearchGroup);
                  return;
                }
                if (principalCanCreate) {
                  void handleCreateGroupAndLink("principal", principalSearch);
                }
              }}
              title={principalCanCreate ? "Criar grupo" : exactPrincipalSearchGroup ? "Selecionar grupo existente" : "Pesquisar grupo"}
            >
              {principalCanCreate ? "Criar" : "Pesquisar"}
            </button>
          </div>
          {principalSearchResults.length ? (
            <div style={S.searchResultListCompact}>
              {principalSearchResults.map((group) => (
                <button
                  key={group.id}
                  type="button"
                  style={group.id === principalGroupId ? S.searchResultBtnOn : S.searchResultBtn}
                  onClick={() => selectPrincipalGroup(group)}
                >
                  <span>{group.name}</span>
                  {group.id === principalGroupId ? <span style={S.resultMiniMeta}>Selecionado</span> : null}
                </button>
              ))}
            </div>
          ) : null}
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Selecionado</div>
          <div style={S.editorValueStrong}>{principalGroup?.name || (principalCanCreate ? principalSearch || "--" : "--")}</div>
        </div>
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorOptionGrid}>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.principalCategorize} onChange={(event) => updateClassificationMeta({ principalCategorize: event.target.checked })} /> Grupo em categoria Outlook</label>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.principalStatusCategorize} onChange={(event) => updateClassificationMeta({ principalStatusEnabled: event.target.checked, principalStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
            </div>
              <StatusLegend />
          </div>
        ) : null}
      </div>
    );
  }

  function renderLabelsEditor() {
    const suggestionChips = suggestedLabelSeeds.map((label) => ({
      key: label,
      label,
      active: selectedLabels.includes(label),
      onClick: () => applySuggestedLabel(label),
    }));
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Etiquetas</div>
        <div style={S.editorLead}>Liga ou desliga apenas as etiquetas relevantes.</div>
        {renderSuggestionTray("labels", "Sugestoes da leitura", suggestionChips)}
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Pesquisar ou criar</div>
          <div style={S.searchInlineRow}>
            <input
              style={S.input}
              value={classificationLabelInput}
              onChange={(event) => setClassificationLabelInput(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  handleClassificationLabelSearchAction();
                }
              }}
              placeholder="Escreve o nome da etiqueta..."
            />
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={handleClassificationLabelSearchAction}
              disabled={!String(classificationLabelInput || "").trim()}
            >
              {classificationLabelCanCreate ? "Criar" : "Ligar"}
            </button>
          </div>
          {filteredClassificationLabels.length && String(classificationLabelInput || "").trim() ? (
            <div style={S.searchResultListCompact}>
              {filteredClassificationLabels.map((label) => (
                <button
                  key={label}
                  type="button"
                  style={selectedLabels.includes(label) ? S.searchResultBtnOn : S.searchResultBtn}
                  onClick={() => {
                    if (selectedLabels.includes(label)) {
                      removeLabel(label);
                    } else {
                      addLabel(label);
                    }
                    setClassificationLabelInput(label);
                  }}
                >
                  <span>{label}</span>
                  {selectedLabels.includes(label) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
                </button>
              ))}
            </div>
          ) : String(classificationLabelInput || "").trim() ? (
            <div style={S.cardMeta}>
              {classificationLabelCanCreate
                ? `Ainda nao existe nenhuma etiqueta com este nome. Usa Criar para adicionar "${String(classificationLabelInput || "").trim()}".`
                : "Etiqueta exata encontrada. Usa Ligar para a associar ou remover."}
            </div>
          ) : null}
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Selecionadas</div>
          <div style={S.chipGridCompact}>
            {selectedLabels.length ? selectedLabels.map((label) => (
              <button key={label} type="button" style={S.groupChipBtnOn} onClick={() => removeLabel(label)}>{label}</button>
            )) : <span style={S.mutedMini}>Sem etiquetas selecionadas.</span>}
          </div>
        </div>
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Estado da etiqueta</span>
                <select
                  style={S.select}
                  value={selectedLabelSharedStatus}
                  onChange={(event) => {
                    const nextValue = String(event.target.value || "").trim() as EmailLabelStatus | "";
                    selectedLabels.forEach((label) => updateLabelDraft(label, {
                      hasStatus: Boolean(nextValue),
                      status: nextValue || undefined,
                    }));
                  }}
                >
                  <option value="">Sem estado</option>
                  {LABEL_STATUS_OPTIONS.map((option) => <option key={option.value} value={option.value}>{option.label}</option>)}
                </select>
              </label>
              <label style={S.compactCheckBoxField}><input type="checkbox" checked={selectedLabels.some((label) => labelDrafts[label]?.categorize === true)} onChange={(event) => selectedLabels.forEach((label) => updateLabelDraft(label, { categorize: event.target.checked }))} /> Etiqueta em categoria Outlook</label>
            </div>
            <label style={S.compactCheckBoxField}><input type="checkbox" checked={selectedLabels.some((label) => labelDrafts[label]?.hasStatus === true)} onChange={(event) => selectedLabels.forEach((label) => updateLabelDraft(label, { hasStatus: event.target.checked, status: event.target.checked ? (labelDrafts[label]?.status || "em_analise") : undefined }))} /> Refletir estado pela cor da categoria</label>
              <StatusLegend />
          </div>
        ) : null}
      </div>
    );
  }

  function renderTicketEditor() {
    const activeList = normalizedTicketSearch ? ticketSearchResults : availableTicketChoices.slice(0, 8);
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Ticket</div>
        <div style={S.editorLead}>Liga um ticket so se houver seguimento operacional.</div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Estado atual</div>
          <div style={S.editorValueStrong}>{selectedTicket?.code || (selectedSeriesId ? "Novo ticket preparado" : "Sem ticket ligado")}</div>
        </div>
        <div style={S.editorSplitRow}>
          <button type="button" style={ticketEditorMode === "existing" ? S.editorModeBtnOn : S.editorModeBtn} onClick={() => setTicketEditorMode("existing")}>Ligar ticket existente</button>
          <button type="button" style={ticketEditorMode === "new" ? S.editorModeBtnOn : S.editorModeBtn} onClick={() => setTicketEditorMode("new")}>Criar novo ticket</button>
        </div>
        {ticketEditorMode === "existing" ? (
          <div style={S.editorBlock}>
            <div style={S.searchInlineRow}>
              <input
                style={S.input}
                value={ticketSearch}
                onChange={(event) => setTicketSearch(event.target.value)}
                onKeyDown={(event) => {
                  if (event.key === "Enter") {
                    event.preventDefault();
                    void handleSearchTickets(undefined, { silent: false });
                  }
                }}
                placeholder="Pesquisar por codigo, titulo ou etiqueta..."
              />
              <button type="button" style={S.secondaryBtn} onClick={() => void handleSearchTickets(undefined, { silent: false })} disabled={!normalizedTicketSearch}>
                Procurar
              </button>
            </div>
            <div style={S.searchResultListCompact}>
              {activeList.length ? activeList.map((ticket) => (
                <button key={ticket.id} type="button" style={ticket.id === selectedTicketId ? S.searchResultBtnOn : S.searchResultBtn} onClick={() => applySuggestedTicket(ticket.id)}>
                  <span>{ticket.code || ticket.title || "Ticket"}</span>
                  {ticket.title && ticket.title !== ticket.code ? <span style={S.resultMiniMeta}>{ticket.title}</span> : null}
                  {ticket.id === selectedTicketId ? <span style={S.resultMiniMeta}>Ligado</span> : null}
                </button>
              )) : (
                <span style={S.mutedMini}>
                  {normalizedTicketSearch
                    ? (ticketSearchBusy ? "A procurar tickets..." : "Nenhum ticket encontrado para esta pesquisa.")
                    : "Sem tickets disponiveis para ligar."}
                </span>
              )}
            </div>
          </div>
        ) : (
          <div style={S.editorBlock}>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Serie</span>
                <select style={S.select} value={selectedSeriesId} onChange={(event) => { setSelectedSeriesId(event.target.value); setSelectionTouched((current) => ({ ...current, ticket: true })); }}>
                  <option value="">Escolher serie...</option>
                  {ticketSeries.map((series) => <option key={series.id} value={series.id}>{series.prefix} · {series.name}</option>)}
                </select>
              </label>
              <label style={S.field}>
                <span style={S.cardMeta}>Titulo</span>
                <input style={S.input} value={createTicketTitle} onChange={(event) => setCreateTicketTitle(event.target.value)} placeholder="Titulo do novo ticket" />
              </label>
            </div>
          </div>
        )}
        {classificationLayoutMode === "advanced" ? (
          <div style={S.editorBlock}>
            <div style={S.editorBlockTitle}>Opcoes avancadas</div>
            <div style={S.editorAdvancedFieldGrid}>
              <label style={S.field}>
                <span style={S.cardMeta}>Estado do ticket</span>
                <select style={S.select} value={ticketStatusDraft} onChange={(event) => setTicketStatusDraft(event.target.value)}>
                  {TICKET_STATUS_OPTIONS.map((option) => <option key={option.value || "none"} value={option.value}>{option.label}</option>)}
                </select>
              </label>
              <label style={S.compactCheckBoxField}><input type="checkbox" checked={Boolean(selectedTicketId || selectedSeriesId)} onChange={(event) => { if (!event.target.checked) clearTicketSelection(); }} /> Ticket em categoria Outlook</label>
            </div>
            <label style={S.compactCheckBoxField}><input type="checkbox" checked={classificationMetaDraft.ticketStatusCategorize} onChange={(event) => updateClassificationMeta({ ticketStatusEnabled: event.target.checked, ticketStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
              <StatusLegend />
          </div>
        ) : null}
      </div>
    );
  }

  function renderReferencesEditor() {
    if (classificationLayoutMode !== "advanced") {
      return (
        <div style={S.editorPanelStack}>
          <div style={S.editorModeKicker}>Referencias</div>
          <div style={S.editorLead}>As referencias so aparecem no modo avancado.</div>
        </div>
      );
    }
    return (
      <div style={S.editorPanelStack}>
        <div style={S.editorModeKicker}>Referencias</div>
        <div style={S.editorLead}>Liga este caso a outros dossiers apenas quando houver ligacao estrutural real.</div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Ligadas</div>
          <div style={S.chipGridCompact}>
            {referenceGroups.length ? referenceGroups.map((group) => (
              <button key={group.id} type="button" style={S.groupChipBtnOn} onClick={() => toggleReferenceGroup(group.id)}>
                {group.name || group.id}
              </button>
            )) : <span style={S.mutedMini}>Sem referencias ligadas.</span>}
          </div>
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Pesquisar outro dossier</div>
          <div style={S.searchInlineRow}>
            <input
              style={S.input}
              value={referenceSearch}
              onChange={(event) => setReferenceSearchValue(event.target.value)}
              onKeyDown={(event) => {
                if (event.key === "Enter") {
                  event.preventDefault();
                  if (referenceCanCreate) {
                    void handleCreateGroupAndLink("referencia", referenceSearch);
                  } else if (exactReferenceSearchGroup) {
                    toggleReferenceGroup(exactReferenceSearchGroup.id);
                    setReferenceSearchValue(exactReferenceSearchGroup.name);
                  }
                }
              }}
              placeholder="Escreve para pesquisar..."
            />
            <button
              type="button"
              style={S.secondaryBtn}
              onClick={() => {
                if (referenceCanCreate) {
                  void handleCreateGroupAndLink("referencia", referenceSearch);
                  return;
                }
                if (exactReferenceSearchGroup) {
                  toggleReferenceGroup(exactReferenceSearchGroup.id);
                  setReferenceSearchValue(exactReferenceSearchGroup.name);
                }
              }}
              title={referenceCanCreate ? "Criar referencia" : exactReferenceSearchGroup ? "Ligar ou desligar referencia existente" : "Pesquisar referencia"}
            >
              {referenceCanCreate ? "Criar" : "Procurar"}
            </button>
          </div>
          {!referenceSearchResults.length && String(referenceSearch || "").trim() ? (
            <div style={S.cardMeta}>
              {referenceCanCreate
                ? `Ainda nao existe nenhum dossier com este nome. Usa Procurar para criar "${String(referenceSearch || "").trim()}".`
                : "Referencia exata encontrada. Usa Procurar para a ligar ou desligar."}
            </div>
          ) : null}
          {referenceSearchResults.length ? (
            <div style={S.searchResultListCompact}>
              {referenceSearchResults.map((group) => (
                <button key={group.id} type="button" style={referenceGroupIds.includes(group.id) ? S.searchResultBtnOn : S.searchResultBtn} onClick={() => toggleReferenceGroup(group.id)}>
                  <span>{group.name}</span>
                  {referenceGroupIds.includes(group.id) ? <span style={S.resultMiniMeta}>Ligada</span> : null}
                </button>
              ))}
            </div>
          ) : null}
        </div>
        <div style={S.editorBlock}>
          <div style={S.editorBlockTitle}>Opcoes avancadas</div>
          <div style={S.editorOptionStackLoose}>
            <div style={S.editorOptionGrid}>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.referenceCategorize} onChange={(event) => updateClassificationMeta({ referenceCategorize: event.target.checked })} /> Referencia em categoria Outlook</label>
              <label style={S.compactCheck}><input type="checkbox" checked={classificationMetaDraft.referenceStatusCategorize} onChange={(event) => updateClassificationMeta({ referenceStatusEnabled: event.target.checked, referenceStatusCategorize: event.target.checked })} /> Refletir estado pela cor da categoria</label>
            </div>
            <div style={S.editorLegendWrap}>
              <StatusLegend />
            </div>
          </div>
        </div>
      </div>
    );
  }

  // Final content layout branch
  if (classificationFocus === "principal") return renderPrincipalEditor();
  if (classificationFocus === "labels") return renderLabelsEditor();
  if (classificationFocus === "ticket") return renderTicketEditor();
  return renderReferencesEditor();
};

const S: Record<string, React.CSSProperties> = {
  editorPanelStack: { display: "flex", flexDirection: "column", gap: 16 },
  editorModeKicker: { fontSize: 10, fontWeight: 700, color: "var(--skin-accent-main)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 },
  editorLead: { fontSize: 12, color: "var(--skin-text-muted)", marginBottom: 8, lineHeight: "1.4" },
  editorBlock: { display: "flex", flexDirection: "column", gap: 8 },
  editorBlockHeader: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  editorBlockTitle: { fontSize: 11, fontWeight: 700, color: "var(--skin-text-muted)", textTransform: "uppercase" },
  chevronBtn: { background: "none", border: "none", width: 16, height: 16, cursor: "pointer", color: "var(--skin-text-muted)", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center" },
  chipGridCompact: { display: "flex", flexWrap: "wrap", gap: 6 },
  miniChip: { padding: "4px 8px", fontSize: 10, borderRadius: 4, background: "var(--skin-bg-muted)", color: "var(--skin-text-muted)", border: "1px solid var(--skin-border-main)", cursor: "pointer", whiteSpace: "nowrap" },
  miniChipOn: { padding: "4px 8px", fontSize: 10, borderRadius: 4, background: "var(--skin-bg-active)", color: "var(--skin-accent-main)", border: "1px solid var(--skin-accent-main)", cursor: "pointer", whiteSpace: "nowrap" },
  editorExpandableOpen: { background: "var(--skin-bg-muted)", borderRadius: 8, padding: 8, border: "1px solid var(--skin-border-main)" },
  editorExpandableScroll: { display: "flex", flexWrap: "wrap", gap: 6, maxHeight: 120, overflowY: "auto" },
  searchInlineRow: { display: "flex", gap: 8 },
  input: { flex: 1, height: 28, fontSize: 11, padding: "0 8px", borderRadius: 4, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-input)", color: "var(--skin-text-main)" },
  secondaryBtn: { display: "flex", alignItems: "center", gap: 6, padding: "0 10px", height: 28, fontSize: 11, fontWeight: 600, borderRadius: 4, background: "var(--skin-bg-main)", color: "var(--skin-text-main)", border: "1px solid var(--skin-border-main)", cursor: "pointer" },
  searchResultListCompact: { display: "flex", flexDirection: "column", gap: 1, maxHeight: 160, overflowY: "auto", border: "1px solid var(--skin-border-main)", borderRadius: 4, background: "var(--skin-bg-muted)", padding: 1 },
  searchResultBtn: { display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", padding: "6px 8px", border: "none", borderBottom: "1px solid var(--skin-border-main)", background: "var(--skin-bg-main)", color: "var(--skin-text-main)", fontSize: 11, textAlign: "left", cursor: "pointer" },
  searchResultBtnOn: { display: "flex", justifyContent: "space-between", alignItems: "center", width: "100%", padding: "6px 8px", border: "none", borderBottom: "1px solid var(--skin-border-main)", background: "var(--skin-bg-active)", color: "var(--skin-accent-main)", fontSize: 11, fontWeight: 600, textAlign: "left", cursor: "pointer" },
  resultMiniMeta: { fontSize: 9, color: "var(--skin-text-muted)", fontWeight: 400 },
  editorValueStrong: { fontSize: 14, fontWeight: 700, color: "var(--skin-text-main)", padding: "4px 0" },
  editorOptionGrid: { display: "flex", flexDirection: "column", gap: 6 },
  compactCheck: { display: "flex", alignItems: "center", gap: 8, fontSize: 11, color: "var(--skin-text-main)", cursor: "pointer" },
  editorAdvancedFieldGrid: { display: "flex", flexDirection: "column", gap: 12 },
  field: { display: "flex", flexDirection: "column", gap: 4 },
  cardMeta: { fontSize: 10, color: "var(--skin-text-muted)", lineHeight: "1.4" },
  select: { height: 28, fontSize: 11, padding: "0 4px", borderRadius: 4, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-input)", color: "var(--skin-text-main)" },
  compactCheckBoxField: { display: "flex", alignItems: "center", gap: 8, fontSize: 11, color: "var(--skin-text-main)", cursor: "pointer" },
  groupChipBtn: { padding: "4px 8px", fontSize: 11, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-main)", color: "var(--skin-text-muted)", borderRadius: 4, cursor: "pointer" },
  groupChipBtnOn: { padding: "4px 8px", fontSize: 11, border: "1px solid var(--skin-accent-main)", background: "var(--skin-bg-active)", color: "var(--skin-accent-main)", borderRadius: 4, cursor: "pointer", fontWeight: 600 },
  editorSplitRow: { display: "flex", gap: 1, background: "var(--skin-border-main)", borderRadius: 6, overflow: "hidden", border: "1px solid var(--skin-border-main)" },
  editorModeBtn: { flex: 1, padding: "8px 4px", fontSize: 11, border: "none", background: "var(--skin-bg-main)", color: "var(--skin-text-muted)", cursor: "pointer", fontWeight: 500 },
  editorModeBtnOn: { flex: 1, padding: "8px 4px", fontSize: 11, border: "none", background: "var(--skin-bg-active)", color: "var(--skin-accent-main)", cursor: "pointer", fontWeight: 600 },
  mutedMini: { fontSize: 10, color: "var(--skin-text-muted)", fontStyle: "italic" },
  editorOptionStackLoose: { display: "flex", flexDirection: "column", gap: 12 },
  editorLegendWrap: { marginTop: 4 },
};

export default ClassificationEditor;
