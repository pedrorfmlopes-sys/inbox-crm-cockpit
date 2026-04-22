import React, { useCallback, useEffect, useMemo, useState } from "react";
import { HelpHint } from "@/ui/HelpHint";
import * as Icons from "@/ui/icons";
import {
  cleanupIntermediateCases,
  migrateIntermediateCaseNamespace,
} from "../storage/intermediateCaseMaintenance";
import {
  normalizeGroupsTabSettings,
  type GroupsSettingsAttachmentStrategy,
  type GroupsSettingsFrequency,
  type GroupsSettingsMigrationMode,
  type GroupsSettingsStorageMode,
  type GroupsTabSettings,
} from "./groupsTabSettings";

type GroupsSettingsSection =
  | "general"
  | "intermediate_storage"
  | "attachments"
  | "cleanup"
  | "warnings"
  | "migration"
  | "maintenance"
  | "explore"
  | "about";

type Props = {
  open: boolean;
  value: GroupsTabSettings | null;
  onClose: () => void;
  onSave: (draft: GroupsTabSettings) => Promise<void> | void;
};

type SectionEntry = {
  id: GroupsSettingsSection;
  label: string;
  icon: React.ReactNode;
};

type PathEditorState = {
  field: "baseFolderPath" | "migrationTarget";
  label: string;
  value: string;
};

const SECTION_ENTRIES: SectionEntry[] = [
  { id: "general", label: "General", icon: <Icons.Settings size={12} /> },
  { id: "intermediate_storage", label: "Storage intermedio", icon: <Icons.Database size={12} /> },
  { id: "attachments", label: "Anexos", icon: <Icons.Paperclip size={12} /> },
  { id: "cleanup", label: "Limpeza", icon: <Icons.RefreshCw size={12} /> },
  { id: "warnings", label: "Avisos", icon: <Icons.AlertCircle size={12} /> },
  { id: "migration", label: "Migracao", icon: <Icons.Upload size={12} /> },
  { id: "maintenance", label: "Manutencao", icon: <Icons.Trash size={12} /> },
  { id: "explore", label: "Explorar", icon: <Icons.Search size={12} /> },
  { id: "about", label: "Sobre", icon: <Icons.MessageSquare size={12} /> },
];

const STORAGE_MODE_OPTIONS: Array<{ value: GroupsSettingsStorageMode; label: string }> = [
  { value: "local_indexeddb", label: "Storage local do add-in (IndexedDB)" },
  { value: "disabled", label: "Desativado" },
];

const ATTACHMENT_STRATEGY_LABELS: Record<GroupsSettingsAttachmentStrategy, string> = {
  server: "Server metadata first",
  outside: "Fora do server quando suportado",
  by_size: "Metadata first com decisao por tamanho",
};

const FREQUENCY_LABELS: Record<GroupsSettingsFrequency, string> = {
  manual: "Manual",
  daily: "Diaria",
  weekly: "Semanal",
};

const MIGRATION_MODE_LABELS: Record<GroupsSettingsMigrationMode, string> = {
  always_ask: "Perguntar sempre",
  move: "Mover quando confirmado",
  copy: "Copiar quando confirmado",
};

function buildDraft(value: GroupsTabSettings | null | undefined): GroupsTabSettings {
  return normalizeGroupsTabSettings(value || null);
}

function SectionButton({
  entry,
  active,
  compact = false,
  onClick,
}: {
  entry: SectionEntry;
  active: boolean;
  compact?: boolean;
  onClick: () => void;
}) {
  return (
    <button
      type="button"
      style={
        active
          ? compact
            ? { ...S.sidebarButtonActive, ...S.sidebarButtonCompactActive }
            : S.sidebarButtonActive
          : compact
            ? { ...S.sidebarButton, ...S.sidebarButtonCompact }
            : S.sidebarButton
      }
      onClick={onClick}
    >
      <span style={S.sidebarIcon}>{entry.icon}</span>
      <span style={S.sidebarLabel}>{entry.label}</span>
    </button>
  );
}

function SettingHint({ text }: { text: string }) {
  return <HelpHint text={text} title="Ajuda do campo" />;
}

function SectionShell({
  title,
  subtitle,
  children,
}: {
  title: string;
  subtitle: string;
  children: React.ReactNode;
}) {
  return (
    <div style={S.sectionShell}>
      <div style={S.sectionHeader}>
        <div style={S.sectionTitle}>{title}</div>
        <div style={S.sectionSubtitle}>{subtitle}</div>
      </div>
      <div style={S.sectionBody}>{children}</div>
    </div>
  );
}

function InfoBlock({
  tone = "neutral",
  title,
  children,
}: {
  tone?: "neutral" | "warning";
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div style={tone === "warning" ? S.infoBlockWarning : S.infoBlock}>
      <div style={S.infoBlockTitle}>{title}</div>
      <div style={S.infoBlockBody}>{children}</div>
    </div>
  );
}

function FieldRow({
  label,
  hint,
  disabled = false,
  children,
}: {
  label: string;
  hint?: string;
  disabled?: boolean;
  children: React.ReactNode;
}) {
  return (
    <label style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <div style={S.rowControl}>{children}</div>
    </label>
  );
}

function ToggleRow({
  label,
  hint,
  checked,
  disabled = false,
  onChange,
}: {
  label: string;
  hint?: string;
  checked: boolean;
  disabled?: boolean;
  onChange: (next: boolean) => void;
}) {
  return (
    <div style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <button
        type="button"
        style={disabled ? S.toggleDisabled : checked ? S.toggleOn : S.toggleOff}
        onClick={() => !disabled && onChange(!checked)}
        aria-pressed={checked}
        disabled={disabled}
      >
        <span style={S.toggleThumb} />
      </button>
    </div>
  );
}

function ActionButton({
  label,
  tone = "neutral",
  disabled = false,
  onClick,
}: {
  label: string;
  tone?: "neutral" | "danger";
  disabled?: boolean;
  onClick?: () => void;
}) {
  return (
    <button
      type="button"
      style={disabled ? S.actionButtonDisabled : tone === "danger" ? S.actionButtonDanger : S.actionButton}
      onClick={disabled ? undefined : onClick}
      disabled={disabled}
    >
      {label}
    </button>
  );
}

function PathFieldRow({
  label,
  hint,
  value,
  chooseLabel = "Definir",
  showValidate = false,
  showOpen = false,
  disabled = false,
  onChoose,
}: {
  label: string;
  hint?: string;
  value: string;
  chooseLabel?: string;
  showValidate?: boolean;
  showOpen?: boolean;
  disabled?: boolean;
  onChoose?: () => void;
}) {
  const displayValue = value || "Nao definido nesta fase";
  return (
    <div style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <div style={S.pathControl}>
        <div style={S.pathValue} title={displayValue}>
          {displayValue}
        </div>
        <div style={S.pathActions}>
          <ActionButton label={chooseLabel} onClick={onChoose} disabled={disabled} />
          {showValidate ? <ActionButton label="Validar" disabled /> : null}
          {showOpen ? <ActionButton label="Abrir pasta" disabled /> : null}
        </div>
      </div>
    </div>
  );
}

function ActionRow({
  label,
  actionLabel,
  tone = "neutral",
  disabled = false,
}: {
  label: string;
  actionLabel: string;
  tone?: "neutral" | "danger";
  disabled?: boolean;
}) {
  return (
    <div style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
      </div>
      <div style={S.rowControl}>
        <div style={S.pathActions}>
          <ActionButton label={actionLabel} tone={tone} disabled={disabled} />
        </div>
      </div>
    </div>
  );
}

function readOnlyBoolean(value: boolean): string {
  return value ? "Ativo" : "Desativado";
}

export function GroupsSettingsPanel({ open, value, onClose, onSave }: Props): JSX.Element | null {
  const [activeSection, setActiveSection] = useState<GroupsSettingsSection>("general");
  const [draft, setDraft] = useState<GroupsTabSettings>(() => buildDraft(value));
  const [isSaving, setIsSaving] = useState(false);
  const [maintenanceMessage, setMaintenanceMessage] = useState("");
  const [maintenanceError, setMaintenanceError] = useState("");
  const [isCompactLayout, setIsCompactLayout] = useState(false);
  const [pathEditor, setPathEditor] = useState<PathEditorState | null>(null);

  const applyDraftPatch = (patch: Partial<GroupsTabSettings>) => {
    setDraft((current) => normalizeGroupsTabSettings({ ...current, ...patch }));
  };

  const openPathEditor = useCallback(
    (field: PathEditorState["field"], label: string) => {
      setPathEditor({
        field,
        label,
        value: String(draft[field] || "").trim(),
      });
    },
    [draft]
  );

  const closePathEditor = useCallback(() => {
    setPathEditor(null);
  }, []);

  const applyPathEditor = useCallback(() => {
    if (!pathEditor) return;
    applyDraftPatch({
      [pathEditor.field]: String(pathEditor.value || "").trim(),
    } as Partial<GroupsTabSettings>);
    setPathEditor(null);
  }, [pathEditor]);

  useEffect(() => {
    if (!open) return;
    setDraft(buildDraft(value));
    setActiveSection("general");
    setIsSaving(false);
    setMaintenanceMessage("");
    setMaintenanceError("");
    setPathEditor(null);
  }, [open, value]);

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape" && !isSaving) onClose();
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isSaving, onClose, open]);

  useEffect(() => {
    if (!open || typeof window === "undefined") return;
    const syncLayoutMode = () => setIsCompactLayout(window.innerWidth <= 760);
    syncLayoutMode();
    window.addEventListener("resize", syncLayoutMode);
    return () => window.removeEventListener("resize", syncLayoutMode);
  }, [open]);

  const handleSave = async () => {
    setIsSaving(true);
    try {
      await onSave(normalizeGroupsTabSettings(draft));
    } finally {
      setIsSaving(false);
    }
  };

  const handleCleanupNow = useCallback(async () => {
    setIsSaving(true);
    setMaintenanceMessage("");
    setMaintenanceError("");
    try {
      const result = await cleanupIntermediateCases(draft);
      setMaintenanceMessage(
        result.inspectedCases
          ? `Limpeza concluida: ${result.deletedCases} caso(s) apagado(s), ${result.skippedMixedCases} misto(s) preservado(s).`
          : "Nao havia casos intermédios persistidos para limpar neste namespace."
      );
    } catch (error) {
      setMaintenanceError(error instanceof Error ? error.message : "Nao foi possivel executar a limpeza real do intermédio.");
    } finally {
      setIsSaving(false);
    }
  }, [draft]);

  const handleMigrationNow = useCallback(async () => {
    setIsSaving(true);
    setMaintenanceMessage("");
    setMaintenanceError("");
    try {
      const sourceNamespace = String(draft.baseFolderPath || "").trim();
      const targetNamespace = String(draft.migrationTarget || "").trim();
      if (!sourceNamespace || !targetNamespace) {
        throw new Error("Define namespaces de origem e destino antes de correr a migracao real do intermédio.");
      }
      const mode = draft.migrationMode === "move" ? "move" : "copy";
      const result = await migrateIntermediateCaseNamespace({
        sourceNamespace,
        targetNamespace,
        mode,
      });
      const nextDraft = normalizeGroupsTabSettings({
        ...draft,
        baseFolderPath: targetNamespace,
      });
      setDraft(nextDraft);
      await onSave(nextDraft);
      setMaintenanceMessage(
        `Migracao ${mode === "copy" ? "por copia" : "por movimento"} concluida: ${result.migratedCases} caso(s), ${result.copiedAttachments} anexo(s) copiado(s).`
      );
    } catch (error) {
      setMaintenanceError(error instanceof Error ? error.message : "Nao foi possivel executar a migracao real do intermédio.");
    } finally {
      setIsSaving(false);
    }
  }, [draft, onSave]);

  const content = useMemo(() => {
    if (activeSection === "general") {
      return (
        <SectionShell title="General" subtitle="Estado base da aba Groups nesta fase executavel.">
          <ToggleRow
            label="Aba Groups ativa"
            hint="Este switch continua real e controla o gating local da aba."
            checked={draft.groupsTabEnabled}
            onChange={(next) => applyDraftPatch({ groupsTabEnabled: next })}
          />
          <FieldRow label="Versao desta fase" hint="Referencia curta do bloco funcional em uso.">
            <div style={S.inlineValue}>{draft.groupsVersion}</div>
          </FieldRow>
          <FieldRow label="Diagnostico rapido" hint="Resumo honesto do modo de gravacao atual.">
            <div style={S.inlineValue}>{draft.quickDiagnostic}</div>
          </FieldRow>
          <InfoBlock title="Fronteira executavel desta fase">
            <ul style={S.infoList}>
              <li>Intermedio: IndexedDB local do add-in quando existe namespace; memoria quando nao existe.</li>
              <li>Final: persistencia classificada via pipeline atual da app (`/api/links/*`).</li>
              <li>Sessao/cache: seeds locais e rascunhos temporarios nao contam como gravacao final.</li>
            </ul>
          </InfoBlock>
        </SectionShell>
      );
    }

    if (activeSection === "intermediate_storage") {
      return (
        <SectionShell
          title="Storage intermedio"
          subtitle="Nesta fase, o intermedio e apenas a ponte local entre Preparar e Classificar."
        >
          <FieldRow label="Modo de storage" hint="So existe storage local do add-in ou modo desligado nesta fase.">
            <select
              style={S.select}
              value={draft.storageMode}
              onChange={(event) => applyDraftPatch({ storageMode: event.target.value as GroupsSettingsStorageMode })}
            >
              {STORAGE_MODE_OPTIONS.map((option) => (
                <option key={option.value} value={option.value}>
                  {option.label}
                </option>
              ))}
            </select>
          </FieldRow>
          <PathFieldRow
            label="Namespace persistente"
            hint="Chave logica usada para namespacing do IndexedDB local. Nao e uma pasta real nem um path web."
            value={draft.baseFolderPath}
            chooseLabel="Definir namespace"
            showValidate={false}
            showOpen={false}
            onChoose={() => openPathEditor("baseFolderPath", "Namespace do storage intermedio")}
          />
          <FieldRow label="Estado" hint="Resumo do modo intermedio realmente executavel.">
            <div style={S.inlineValue}>{draft.locationStatus}</div>
          </FieldRow>
          <InfoBlock title="O que grava de verdade">
            <ul style={S.infoList}>
              <li>Com namespace: `IntermediateCase` em IndexedDB local do host.</li>
              <li>Sem namespace: fallback em memoria, sem persistencia local duravel.</li>
              <li>Desativado: sem storage intermedio e sem promessa de retoma local.</li>
            </ul>
          </InfoBlock>
          <InfoBlock tone="warning" title="Preferencias ainda nao executaveis">
            <div style={S.inlineValue}>
              Os toggles abaixo continuam apenas como shell herdada. Nao abrem validacao real de pasta,
              migracao, nem bridge para OneDrive/SharePoint nesta fase.
            </div>
          </InfoBlock>
          <ToggleRow
            label="Criar caso automaticamente ao abrir email novo"
            hint="Mantido como preferencia herdada; ainda nao muda o runtime de ponta a ponta nesta fase."
            checked={draft.autoCreateCaseOnNewEmail}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Reabrir caso existente"
            hint="A reabertura real existe, mas este toggle ainda nao e o contrato executavel que a controla."
            checked={draft.reopenExistingCase}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Recriar copia intermedia a partir do servidor"
            hint="Preferencia reservada; sem motor independente adicional nesta fase."
            checked={draft.recreateIntermediateCopy}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Validar localizacao ao abrir"
            hint="Nao existe validacao real de pasta porque o storage intermedio nao usa pasta fisica."
            checked={draft.validateLocationOnOpen}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Bloquear a aba se a localizacao falhar"
            hint="Nao existe localizacao fisica a validar; o bloqueio real continua a ser por modo/namespace."
            checked={draft.blockTabIfUnavailable}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Avisar indisponibilidade"
            hint="Preferencia herdada, sem motor de monitorizacao real nesta fase."
            checked={draft.warnIfUnavailable}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Revalidar automaticamente"
            hint="Nao existe ciclo automatico real de revalidacao nesta fase."
            checked={draft.autoRetryValidation}
            disabled
            onChange={() => undefined}
          />
        </SectionShell>
      );
    }

    if (activeSection === "attachments") {
      return (
        <SectionShell
          title="Anexos"
          subtitle="Politica executavel desta fase: metadata sempre, binario real apenas onde o provider atual suporta."
        >
          <FieldRow label="Estrategia guardada" hint="Estado atual da shell de anexos desta aba.">
            <div style={S.inlineValue}>{ATTACHMENT_STRATEGY_LABELS[draft.attachmentStrategy]}</div>
          </FieldRow>
          <FieldRow label="Limite server (MB)" hint="Limiar hoje usado para decisoes best-effort no pipeline atual.">
            <div style={S.inlineValue}>{draft.attachmentServerLimitMb} MB</div>
          </FieldRow>
          <FieldRow label="Limite intermadio (MB)" hint="Referencia guardada para a shell; nao representa um provider novo.">
            <div style={S.inlineValue}>{draft.attachmentIntermediateLimitMb} MB</div>
          </FieldRow>
          <FieldRow label="Metadata no server" hint="Este comportamento continua real na persistencia final atual.">
            <div style={S.inlineValue}>{readOnlyBoolean(draft.showAttachmentMetadataOnServer)}</div>
          </FieldRow>
          <InfoBlock title="Politica executavel desta fase">
            <ul style={S.infoList}>
              <li>Metadata de anexo sobe sempre quando o payload final inclui o anexo.</li>
              <li>Binario real so e tentado em `cloud` ou em caminho local/sincronizado realmente acessivel.</li>
              <li>`replaceAttachments: false` preserva anexos anteriores quando o payload e parcial.</li>
              <li>Sem provider/path real, fica referencia e metadata; nao ha promessa falsa de escrita binaria.</li>
            </ul>
          </InfoBlock>
          <InfoBlock tone="warning" title="Opcoes ainda nao executaveis nesta shell">
            <div style={S.inlineValue}>
              Pasta externa, preview imediato e switches detalhados abaixo continuam como configuracao herdada.
              O destino final efetivo continua a ser decidido pelo pipeline atual da app.
            </div>
          </InfoBlock>
          <ToggleRow
            label="Guardar anexos no servidor"
            hint="Preferencia herdada; o comportamento real e decidido pelo provider final atual."
            checked={draft.saveAttachmentsOnServer}
            disabled
            onChange={() => undefined}
          />
          <ToggleRow
            label="Guardar anexos fora do servidor"
            hint="Preferencia herdada; nao ativa por si so um destino novo."
            checked={draft.saveAttachmentsOutsideServer}
            disabled
            onChange={() => undefined}
          />
          <PathFieldRow
            label="Destino externo herdado"
            hint="Nao e um destino oficialmente suportado por esta shell nesta fase."
            value={draft.externalAttachmentFolder}
            chooseLabel="Indisponivel"
            disabled
            onChoose={() => undefined}
          />
          <ToggleRow
            label="Preview imediato"
            hint="Abertura/preview continua dependente do host e do storage realmente disponivel."
            checked={draft.requireImmediatePreview}
            disabled
            onChange={() => undefined}
          />
        </SectionShell>
      );
    }

    if (activeSection === "cleanup") {
      return (
        <SectionShell title="Limpeza" subtitle="A shell existe, mas nao abre limpeza real nesta fase.">
          <InfoBlock tone="warning" title="Indisponivel nesta fase">
            <div style={S.inlineValue}>
              A limpeza automatica total continua fora desta ronda, mas a limpeza manual real do
              `IntermediateCase` ja esta disponivel para o namespace atual.
            </div>
          </InfoBlock>
          <FieldRow label="Dias para aviso de caso misto">
            <div style={S.inlineValue}>{draft.mixedCaseWarningDays} dias</div>
          </FieldRow>
          <FieldRow label="Dias para aviso de abandono local">
            <div style={S.inlineValue}>{draft.localAbandonedWarningDays} dias</div>
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso fechado">
            <div style={S.inlineValue}>{draft.cleanupClosedCaseDays} dias</div>
          </FieldRow>
          <FieldRow label="Dias para limpeza de abandono">
            <div style={S.inlineValue}>{draft.cleanupAbandonedCaseDays} dias</div>
          </FieldRow>
          <FieldRow label="Frequencia">
            <div style={S.inlineValue}>{FREQUENCY_LABELS[draft.cleanupFrequency]}</div>
          </FieldRow>
          <ToggleRow
            label="Nunca apagar casos mistos em silencio"
            checked={draft.neverDeleteMixedSilently}
            onChange={(next) => applyDraftPatch({ neverDeleteMixedSilently: next })}
          />
          <div style={S.pathActions}>
            <ActionButton
              label={isSaving ? "A executar" : "Limpar agora"}
              disabled={isSaving || !draft.baseFolderPath || draft.storageMode === "disabled"}
              tone="danger"
              onClick={() => void handleCleanupNow()}
            />
          </div>
        </SectionShell>
      );
    }

    if (activeSection === "warnings") {
      return (
        <SectionShell title="Avisos" subtitle="Avisos guardados como shell; sem motor novo nesta fase.">
          <InfoBlock title="Estado atual">
            <ul style={S.infoList}>
              <li>Os avisos executaveis desta frente continuam a ser locais e leves no cockpit.</li>
              <li>Esta shell nao liga monitorizacao separada, tarefas futuras ou processos recorrentes.</li>
            </ul>
          </InfoBlock>
          <ToggleRow label="Avisar emails por classificar" checked={draft.warnUnclassifiedEmails} disabled onChange={() => undefined} />
          <ToggleRow label="Avisar casos mistos" checked={draft.warnMixedCases} disabled onChange={() => undefined} />
          <FieldRow label="Frequencia">
            <div style={S.inlineValue}>{FREQUENCY_LABELS[draft.warningFrequency]}</div>
          </FieldRow>
          <ToggleRow label="Preparar ponte futura para tarefas" checked={draft.prepareTasksBridge} disabled onChange={() => undefined} />
        </SectionShell>
      );
    }

    if (activeSection === "migration") {
      return (
        <SectionShell title="Migracao" subtitle="Migracao executavel do intermédio por namespace, sem fingir migracao total do storage final.">
          <InfoBlock tone="warning" title="Migracao real do intermédio">
            <div style={S.inlineValue}>
              Nesta ronda a migracao real passa a cobrir o `IntermediateCase` entre namespaces de
              IndexedDB. Migracao historica do storage final continua fora desta shell.
            </div>
          </InfoBlock>
          <PathFieldRow
            label="Destino guardado"
            hint="Namespace alvo para copiar/mover os casos intermédios persistidos."
            value={draft.migrationTarget}
            chooseLabel="Definir destino"
            onChoose={() => openPathEditor("migrationTarget", "Namespace de destino da migracao")}
          />
          <FieldRow label="Modo guardado">
            <select
              style={S.select}
              value={draft.migrationMode}
              onChange={(event) => applyDraftPatch({ migrationMode: event.target.value as GroupsSettingsMigrationMode })}
            >
              {Object.entries(MIGRATION_MODE_LABELS).map(([value, label]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
          </FieldRow>
          <ToggleRow
            label="Permitir mover dados existentes"
            checked={draft.allowMoveExistingData}
            onChange={(next) => applyDraftPatch({ allowMoveExistingData: next })}
          />
          <FieldRow label="Seguranca estrita">
            <div style={S.inlineValue}>{readOnlyBoolean(draft.strictMigrationSafety)}</div>
          </FieldRow>
          <ToggleRow
            label="Fundir dados existentes"
            checked={draft.mergeExistingData}
            onChange={(next) => applyDraftPatch({ mergeExistingData: next })}
          />
          <div style={S.pathActions}>
            <ActionButton
              label={isSaving ? "A migrar" : "Migrar agora"}
              disabled={isSaving || !draft.baseFolderPath || !draft.migrationTarget}
              onClick={() => void handleMigrationNow()}
            />
          </div>
        </SectionShell>
      );
    }

    if (activeSection === "maintenance") {
      return (
        <SectionShell title="Manutencao" subtitle="Sem ferramentas novas de manutencao nesta fase.">
          <InfoBlock tone="warning" title="Acoes desativadas">
            <div style={S.inlineValue}>
              Limpeza profunda, rebuild, verificacao de blobs e tarefas de manutencao continuam fora desta ronda.
            </div>
          </InfoBlock>
          <ActionRow label="Revalidar storage intermedio" actionLabel="Indisponivel" disabled />
          <ActionRow label="Reconstruir indices locais" actionLabel="Indisponivel" disabled />
          <ActionRow label="Limpar dados orfaos" actionLabel="Indisponivel" tone="danger" disabled />
        </SectionShell>
      );
    }

    if (activeSection === "explore") {
      return (
        <SectionShell title="Explorar" subtitle="Mantido explicitamente fora desta fase.">
          <InfoBlock tone="warning" title="Nao aberto nesta ronda">
            <div style={S.inlineValue}>
              `Explorar` e `Gestor do Grupo` continuam fora de scope. Os campos abaixo ficam apenas como
              shell desativada para nao prometer comportamento que ainda nao existe.
            </div>
          </InfoBlock>
          <ToggleRow label="Servidor como base principal" checked={draft.explorerServerPrimary} disabled onChange={() => undefined} />
          <ToggleRow label="Abrir anexos guardados" checked={draft.explorerOpenStoredAttachments} disabled onChange={() => undefined} />
          <ToggleRow label="Permitir gerar resposta e reenvio" checked={draft.explorerGenerateReply} disabled onChange={() => undefined} />
        </SectionShell>
      );
    }

    return (
      <SectionShell title="Sobre" subtitle="Resumo da politica executavel de gravacao da aba Groups nesta fase.">
        <InfoBlock title="Intermedio">
          <ul style={S.infoList}>
            <li>`IntermediateCase` em IndexedDB local quando existe namespace configurado.</li>
            <li>Fallback em memoria quando nao existe namespace ou quando o modo esta desligado.</li>
            <li>Serve para draft, continuidade de sessao e ponte Preparar - Classificar.</li>
          </ul>
        </InfoBlock>
        <InfoBlock title="Final">
          <ul style={S.infoList}>
            <li>Persistencia final atual via pipeline `/api/links/*` da app.</li>
            <li>Sobem email classificado, memberships, tickets e metadata de anexos.</li>
            <li>Binario real so sobe quando o provider/path da frente atual o suporta de verdade.</li>
          </ul>
        </InfoBlock>
        <InfoBlock title="Fora de scope nesta fase">
          <ul style={S.infoList}>
            <li>OneDrive/SharePoint por URL web como destino final.</li>
            <li>Migracao/limpeza automatica.</li>
            <li>Explorar, Gestor do Grupo e nova superficie funcional.</li>
          </ul>
        </InfoBlock>
      </SectionShell>
    );
  }, [activeSection, draft, handleCleanupNow, handleMigrationNow, isSaving, openPathEditor]);

  if (!open) return null;

  return (
    <div style={S.backdrop} role="presentation" onClick={() => !isSaving && onClose()}>
      <div
        style={isCompactLayout ? { ...S.modal, ...S.modalCompact } : S.modal}
        role="dialog"
        aria-modal="true"
        onClick={(event) => event.stopPropagation()}
      >
        <div style={isCompactLayout ? { ...S.modalHeader, ...S.modalHeaderCompact } : S.modalHeader}>
          <div style={S.modalHeaderText}>
            <div style={S.modalEyebrow}>Groups</div>
            <div style={S.modalTitle}>Settings da aba Groups</div>
          </div>
          <div style={isCompactLayout ? { ...S.modalActions, ...S.modalActionsCompact } : S.modalActions}>
            <button type="button" style={S.headerButtonSecondary} onClick={onClose} disabled={isSaving}>
              Fechar
            </button>
            <button
              type="button"
              style={isSaving ? S.headerButtonPrimaryDisabled : S.headerButtonPrimary}
              onClick={handleSave}
              disabled={isSaving}
            >
              <Icons.Save size={12} />
              {isSaving ? "A guardar" : "Guardar"}
            </button>
          </div>
        </div>
        <div style={isCompactLayout ? { ...S.modalBody, ...S.modalBodyCompact } : S.modalBody}>
          <div style={isCompactLayout ? { ...S.sidebar, ...S.sidebarCompact } : S.sidebar}>
            {SECTION_ENTRIES.map((entry) => (
              <SectionButton
                key={entry.id}
                entry={entry}
                active={entry.id === activeSection}
                compact={isCompactLayout}
                onClick={() => setActiveSection(entry.id)}
              />
            ))}
          </div>
          <div style={isCompactLayout ? { ...S.content, ...S.contentCompact } : S.content}>{content}</div>
        </div>
        {pathEditor ? (
          <div style={S.pathEditorShell}>
            <div style={S.pathEditorHeader}>
              <div style={S.pathEditorTitle}>{pathEditor.label}</div>
              <div style={S.pathEditorSubtitle}>Edicao suportada no taskpane, sem usar APIs modais do browser.</div>
            </div>
            <input
              autoFocus
              type="text"
              style={S.input}
              value={pathEditor.value}
              onChange={(event) =>
                setPathEditor((current) => (current ? { ...current, value: event.target.value } : current))
              }
              placeholder="Introduz o valor pretendido"
            />
            <div style={S.pathEditorActions}>
              <ActionButton label="Cancelar" onClick={closePathEditor} disabled={isSaving} />
              <ActionButton label="Aplicar" onClick={applyPathEditor} disabled={isSaving} />
            </div>
          </div>
        ) : null}
        {maintenanceMessage ? (
          <div style={{ ...S.infoBlock, margin: "0 10px 10px" }}>
            <div style={S.infoBlockTitle}>Operacao concluida</div>
            <div style={S.infoBlockBody}>{maintenanceMessage}</div>
          </div>
        ) : null}
        {maintenanceError ? (
          <div style={{ ...S.infoBlockWarning, margin: "0 10px 10px" }}>
            <div style={S.infoBlockTitle}>Operacao bloqueada</div>
            <div style={S.infoBlockBody}>{maintenanceError}</div>
          </div>
        ) : null}
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  backdrop: {
    position: "fixed",
    inset: 0,
    background: "rgba(15,23,42,0.24)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 18,
    zIndex: 70,
  },
  modal: {
    width: "min(960px, calc(100vw - 24px))",
    height: "min(700px, calc(100vh - 24px))",
    borderRadius: 18,
    overflow: "hidden",
    background: "linear-gradient(180deg,#f8fbff 0%, #f3f7fb 100%)",
    boxShadow: "0 24px 70px rgba(15,23,42,0.22)",
    display: "grid",
    gridTemplateRows: "auto minmax(0, 1fr)",
  },
  modalCompact: {
    width: "min(420px, calc(100vw - 12px))",
    height: "min(760px, calc(100vh - 12px))",
    borderRadius: 16,
  },
  modalHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    padding: "10px 12px",
    borderBottom: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.82)",
  },
  modalHeaderCompact: {
    alignItems: "stretch",
    flexDirection: "column",
  },
  modalHeaderText: {
    display: "grid",
    gap: 1,
  },
  modalEyebrow: {
    fontSize: 8.4,
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.05em",
    color: "#64748b",
  },
  modalTitle: {
    fontSize: 13.4,
    fontWeight: 650,
    color: "#243244",
  },
  modalActions: {
    display: "flex",
    alignItems: "center",
    gap: 6,
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  modalActionsCompact: {
    width: "100%",
    justifyContent: "flex-start",
  },
  headerButtonSecondary: {
    borderRadius: 11,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "#fff",
    color: "#526173",
    padding: "5px 9px",
    fontSize: 9,
    fontWeight: 650,
    cursor: "pointer",
  },
  headerButtonPrimary: {
    borderRadius: 11,
    border: "1px solid rgba(37,99,235,0.18)",
    background: "linear-gradient(180deg,#3b82f6 0%, #2563eb 100%)",
    color: "#fff",
    padding: "5px 9px",
    fontSize: 9,
    fontWeight: 700,
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
  },
  headerButtonPrimaryDisabled: {
    borderRadius: 11,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "rgba(148,163,184,0.28)",
    color: "#fff",
    padding: "5px 9px",
    fontSize: 9,
    fontWeight: 700,
    cursor: "wait",
    display: "inline-flex",
    alignItems: "center",
    gap: 4,
  },
  modalBody: {
    display: "grid",
    gridTemplateColumns: "168px minmax(0, 1fr)",
    minHeight: 0,
  },
  modalBodyCompact: {
    gridTemplateColumns: "minmax(0, 1fr)",
    gridTemplateRows: "auto minmax(0, 1fr)",
  },
  sidebar: {
    display: "grid",
    alignContent: "start",
    gap: 4,
    padding: 10,
    borderRight: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.72)",
    overflowY: "auto",
  },
  sidebarCompact: {
    display: "flex",
    gap: 6,
    padding: "8px 8px 6px",
    borderRight: "none",
    borderBottom: "1px solid rgba(148,163,184,0.16)",
    overflowX: "auto",
    overflowY: "hidden",
  },
  sidebarButton: {
    width: "100%",
    border: "1px solid transparent",
    background: "transparent",
    borderRadius: 12,
    padding: "7px 8px",
    display: "grid",
    gridTemplateColumns: "14px minmax(0, 1fr)",
    gap: 7,
    alignItems: "center",
    cursor: "pointer",
    color: "#526173",
    textAlign: "left",
  },
  sidebarButtonActive: {
    width: "100%",
    border: "1px solid rgba(59,130,246,0.16)",
    background: "rgba(219,234,254,0.58)",
    borderRadius: 12,
    padding: "7px 8px",
    display: "grid",
    gridTemplateColumns: "14px minmax(0, 1fr)",
    gap: 7,
    alignItems: "center",
    cursor: "pointer",
    color: "#1e3a5f",
    textAlign: "left",
  },
  sidebarButtonCompact: {
    width: "auto",
    minWidth: 120,
    padding: "8px 10px",
  },
  sidebarButtonCompactActive: {
    width: "auto",
    minWidth: 120,
    padding: "8px 10px",
  },
  sidebarIcon: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
  },
  sidebarLabel: {
    fontSize: 9.6,
    fontWeight: 600,
    lineHeight: 1.2,
  },
  content: {
    minWidth: 0,
    overflowY: "auto",
    padding: 10,
  },
  contentCompact: {
    padding: 8,
  },
  pathEditorShell: {
    display: "grid",
    gap: 8,
    padding: "10px",
    borderTop: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.9)",
  },
  pathEditorHeader: {
    display: "grid",
    gap: 2,
  },
  pathEditorTitle: {
    fontSize: 10.4,
    fontWeight: 650,
    color: "#243244",
  },
  pathEditorSubtitle: {
    fontSize: 8.8,
    color: "#64748b",
    lineHeight: 1.35,
  },
  pathEditorActions: {
    display: "flex",
    gap: 6,
    flexWrap: "wrap",
    justifyContent: "flex-end",
  },
  sectionShell: {
    display: "grid",
    gap: 10,
  },
  sectionHeader: {
    display: "grid",
    gap: 2,
  },
  sectionTitle: {
    fontSize: 11.4,
    fontWeight: 650,
    color: "#243244",
  },
  sectionSubtitle: {
    fontSize: 9.1,
    color: "#64748b",
    lineHeight: 1.35,
  },
  sectionBody: {
    display: "grid",
    gap: 7,
  },
  row: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 10,
    alignItems: "start",
    padding: "7px 8px",
    borderRadius: 12,
    border: "1px solid rgba(148,163,184,0.14)",
    background: "rgba(255,255,255,0.88)",
  },
  rowDisabled: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(220px, 1fr))",
    gap: 10,
    alignItems: "start",
    padding: "7px 8px",
    borderRadius: 12,
    border: "1px solid rgba(148,163,184,0.14)",
    background: "rgba(241,245,249,0.82)",
    opacity: 0.82,
  },
  rowLabelWrap: {
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    minWidth: 0,
  },
  rowLabel: {
    fontSize: 9.4,
    fontWeight: 600,
    color: "#3a495c",
    lineHeight: 1.3,
  },
  rowControl: {
    minWidth: 0,
  },
  pathControl: {
    display: "grid",
    gap: 6,
    minWidth: 0,
  },
  pathValue: {
    minHeight: 28,
    display: "flex",
    alignItems: "center",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.6,
    color: "#243244",
    whiteSpace: "nowrap",
    overflow: "hidden",
    textOverflow: "ellipsis",
    boxSizing: "border-box",
  },
  pathActions: {
    display: "flex",
    gap: 5,
    flexWrap: "wrap",
    justifyContent: "flex-start",
  },
  input: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.6,
    color: "#243244",
    boxSizing: "border-box",
  },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.6,
    color: "#243244",
    boxSizing: "border-box",
  },
  inlineValue: {
    fontSize: 9.5,
    color: "#334155",
    lineHeight: 1.35,
  },
  toggleOn: {
    width: 31,
    height: 18,
    borderRadius: 999,
    border: "1px solid rgba(34,197,94,0.22)",
    background: "rgba(34,197,94,0.72)",
    padding: 1,
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "flex-end",
    cursor: "pointer",
  },
  toggleOff: {
    width: 31,
    height: 18,
    borderRadius: 999,
    border: "1px solid rgba(239,68,68,0.18)",
    background: "rgba(239,68,68,0.58)",
    padding: 1,
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "flex-start",
    cursor: "pointer",
  },
  toggleDisabled: {
    width: 31,
    height: 18,
    borderRadius: 999,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "rgba(148,163,184,0.34)",
    padding: 1,
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "flex-start",
    cursor: "not-allowed",
  },
  toggleThumb: {
    width: 13,
    height: 13,
    borderRadius: 999,
    background: "#fff",
    boxShadow: "0 1px 2px rgba(15,23,42,0.18)",
  },
  actionButton: {
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "#fff",
    color: "#334155",
    padding: "5px 8px",
    fontSize: 9,
    fontWeight: 650,
    cursor: "pointer",
  },
  actionButtonDanger: {
    borderRadius: 10,
    border: "1px solid rgba(220,38,38,0.16)",
    background: "rgba(254,242,242,0.94)",
    color: "#b91c1c",
    padding: "5px 8px",
    fontSize: 9,
    fontWeight: 650,
    cursor: "pointer",
  },
  actionButtonDisabled: {
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.18)",
    background: "rgba(241,245,249,0.88)",
    color: "#94a3b8",
    padding: "5px 8px",
    fontSize: 9,
    fontWeight: 650,
    cursor: "not-allowed",
  },
  infoBlock: {
    display: "grid",
    gap: 6,
    padding: "9px 10px",
    borderRadius: 12,
    border: "1px solid rgba(59,130,246,0.14)",
    background: "rgba(239,246,255,0.82)",
  },
  infoBlockWarning: {
    display: "grid",
    gap: 6,
    padding: "9px 10px",
    borderRadius: 12,
    border: "1px solid rgba(245,158,11,0.18)",
    background: "rgba(255,251,235,0.94)",
  },
  infoBlockTitle: {
    fontSize: 9.4,
    fontWeight: 700,
    color: "#243244",
  },
  infoBlockBody: {
    fontSize: 9.4,
    color: "#334155",
    lineHeight: 1.4,
  },
  infoList: {
    margin: 0,
    paddingLeft: 16,
    display: "grid",
    gap: 4,
  },
};

export default GroupsSettingsPanel;
