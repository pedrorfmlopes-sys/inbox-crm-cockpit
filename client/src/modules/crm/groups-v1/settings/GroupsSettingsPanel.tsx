import React, { useCallback, useEffect, useMemo, useState } from "react";
import { HelpHint } from "@/ui/HelpHint";
import * as Icons from "@/ui/icons";
import {
  cleanupIntermediateCases,
  migrateIntermediateCaseNamespace,
} from "../storage/intermediateCaseMaintenance";
import { resolveIntermediateCaseStorage } from "../storage/resolveIntermediateCaseStorage";
import {
  normalizeGroupsTabSettings,
  type GroupsSettingsAttachmentStrategy,
  type GroupsSettingsFrequency,
  type GroupsSettingsMigrationMode,
  type GroupsSettingsStorageMode,
  type GroupsTabSettings,
} from "./groupsTabSettings";
import { validateGroupsTabStorageAvailability } from "./groupsTabRuntime";

export type GroupsSettingsSection =
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
  initialSection?: GroupsSettingsSection;
  statusMessage?: string;
  statusTone?: "success" | "error";
};

type SectionEntry = {
  id: GroupsSettingsSection;
  label: string;
  icon: React.ReactNode;
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
  always_ask: "Perguntar sempre (indisponivel nesta fase)",
  move: "Mover quando confirmado",
  copy: "Copiar quando confirmado",
};

function buildDraft(value: GroupsTabSettings | null | undefined): GroupsTabSettings {
  return normalizeGroupsTabSettings(value || null);
}

function SectionButton({
  entry,
  active,
  onClick,
}: {
  entry: SectionEntry;
  active: boolean;
  onClick: () => void;
}) {
  return (
    <button type="button" style={active ? S.sidebarButtonActive : S.sidebarButton} onClick={onClick}>
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

function ActionRow({
  label,
  actionLabel,
  tone = "neutral",
  disabled = false,
  hint,
  onClick,
}: {
  label: string;
  actionLabel: string;
  tone?: "neutral" | "danger";
  disabled?: boolean;
  hint?: string;
  onClick?: () => void;
}) {
  return (
    <div style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <div style={S.pathActions}>
        <ActionButton label={actionLabel} tone={tone} disabled={disabled} onClick={onClick} />
      </div>
    </div>
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
  onChangeText,
  placeholder,
}: {
  label: string;
  hint?: string;
  value: string;
  chooseLabel?: string;
  showValidate?: boolean;
  showOpen?: boolean;
  disabled?: boolean;
  onChoose?: () => void;
  onChangeText?: (next: string) => void;
  placeholder?: string;
}) {
  const displayValue = value || "Nao definido nesta fase";
  const isEditable = typeof onChangeText === "function";
  return (
    <div style={disabled ? S.rowDisabled : S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <div style={S.pathControl}>
        {isEditable ? (
          <input
            style={S.input}
            value={value}
            placeholder={placeholder}
            onChange={(event) => onChangeText?.(event.target.value)}
            disabled={disabled}
          />
        ) : (
          <div style={S.pathValue} title={displayValue}>
            {displayValue}
          </div>
        )}
        <div style={S.pathActions}>
          <ActionButton label={chooseLabel} onClick={onChoose} disabled={disabled} />
          {showValidate ? <ActionButton label="Validar" disabled /> : null}
          {showOpen ? <ActionButton label="Abrir pasta" disabled /> : null}
        </div>
      </div>
    </div>
  );
}

function readOnlyBoolean(value: boolean): string {
  return value ? "Ativo" : "Desativado";
}

export function GroupsSettingsPanel({
  open,
  value,
  onClose,
  onSave,
  initialSection = "general",
  statusMessage = "",
  statusTone = "success",
}: Props): JSX.Element | null {
  const [activeSection, setActiveSection] = useState<GroupsSettingsSection>(initialSection);
  const [draft, setDraft] = useState<GroupsTabSettings>(() => buildDraft(value));
  const [isSaving, setIsSaving] = useState(false);
  const [maintenanceMessage, setMaintenanceMessage] = useState("");
  const [maintenanceError, setMaintenanceError] = useState("");

  const applyDraftPatch = (patch: Partial<GroupsTabSettings>) => {
    setDraft((current) => normalizeGroupsTabSettings({ ...current, ...patch }));
  };

  useEffect(() => {
    if (!open) return;
    setDraft(buildDraft(value));
    setActiveSection(initialSection);
    setIsSaving(false);
    setMaintenanceMessage(statusTone === "success" ? statusMessage : "");
    setMaintenanceError(statusTone === "error" ? statusMessage : "");
  }, [initialSection, open, statusMessage, statusTone, value]);

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape" && !isSaving) onClose();
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [isSaving, onClose, open]);

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
          : "Nao havia casos intermédios persistidos para limpar nesta localização."
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
      if (!targetNamespace) {
        throw new Error("Define um destino intermédio antes de correr a migracao real.");
      }
      if (draft.migrationMode === "always_ask") {
        throw new Error("Escolhe primeiro um modo executavel de migracao: copiar ou mover.");
      }
      const mode = draft.migrationMode === "move" ? "move" : "copy";
      const result = await migrateIntermediateCaseNamespace({
        sourceNamespace,
        targetNamespace,
        mode,
        allowMoveExistingData: draft.allowMoveExistingData,
        mergeExistingData: draft.mergeExistingData,
        strictMigrationSafety: draft.strictMigrationSafety,
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

  const handleRevalidateStorage = useCallback(async () => {
    setIsSaving(true);
    setMaintenanceMessage("");
    setMaintenanceError("");
    try {
      const storage = resolveIntermediateCaseStorage(draft);
      const validation = await validateGroupsTabStorageAvailability({
        settings: draft,
        storage,
      });
      if (!validation.available) {
        throw new Error(validation.reason);
      }
      const summaries = await storage.repository.listCases();
      setMaintenanceMessage(`Storage intermédio validado com sucesso. ${summaries.length} caso(s) visível(eis) nesta localização.`);
    } catch (error) {
      setMaintenanceError(error instanceof Error ? error.message : "Nao foi possivel revalidar o storage intermédio.");
    } finally {
      setIsSaving(false);
    }
  }, [draft]);

  const handleRebuildIndexes = useCallback(async () => {
    setIsSaving(true);
    setMaintenanceMessage("");
    setMaintenanceError("");
    try {
      const storage = resolveIntermediateCaseStorage(draft);
      const summaries = await storage.repository.listCases();
      for (const summary of summaries) {
        const caseValue = await storage.repository.readCase(summary.caseId);
        if (caseValue) {
          await storage.repository.writeCase(caseValue);
        }
      }
      setMaintenanceMessage(
        summaries.length
          ? `Indices locais reconstruidos para ${summaries.length} caso(s).`
          : "Nao existiam casos persistidos para reconstruir nesta localização."
      );
    } catch (error) {
      setMaintenanceError(error instanceof Error ? error.message : "Nao foi possivel reconstruir os indices locais.");
    } finally {
      setIsSaving(false);
    }
  }, [draft]);

  const handleCleanupOrphans = useCallback(async () => {
    setIsSaving(true);
    setMaintenanceMessage("");
    setMaintenanceError("");
    try {
      const storage = resolveIntermediateCaseStorage(draft);
      const summaries = await storage.repository.listCases();
      let deletedCases = 0;
      for (const summary of summaries) {
        if (summary.retentionState !== "local_only") continue;
        await storage.repository.deleteCase(summary.caseId);
        deletedCases += 1;
      }
      setMaintenanceMessage(
        summaries.length
          ? `Verificacao de orfaos concluida. ${deletedCases} caso(s) local_only foram removido(s).`
          : "Nao havia dados locais para verificar nesta manutencao."
      );
    } catch (error) {
      setMaintenanceError(error instanceof Error ? error.message : "Nao foi possivel limpar dados orfaos.");
    } finally {
      setIsSaving(false);
    }
  }, [draft]);

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
              <li>Intermedio: pasta local definida quando existe; add-in local como fallback; memória apenas como último recurso.</li>
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
          subtitle="O intermédio arranca logo na pasta local definida; sem pasta, cai para o storage local do add-in."
        >
          <FieldRow label="Modo de storage" hint="Controla se o intermédio fica ativo; com pasta definida ela passa a ser o destino principal.">
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
            label="Pasta local intermédia"
            hint="Quando definida, o `IntermediateCase` grava logo aqui. Se ficar vazia, o fallback principal passa a ser o storage local do add-in."
            value={draft.baseFolderPath}
            chooseLabel="Limpar"
            placeholder="ex.: C:/dados/grupos/intermedio"
            showValidate={false}
            showOpen={false}
            onChangeText={(nextValue) => applyDraftPatch({ baseFolderPath: nextValue })}
            onChoose={() => applyDraftPatch({ baseFolderPath: "" })}
          />
          <FieldRow label="Estado" hint="Resumo do modo intermedio realmente executavel.">
            <div style={S.inlineValue}>{draft.locationStatus}</div>
          </FieldRow>
          <InfoBlock title="O que grava de verdade">
            <ul style={S.infoList}>
              <li>Com pasta definida: `IntermediateCase` grava logo nessa pasta local e a reabertura lê daí.</li>
              <li>Sem pasta definida: o fallback principal é o storage local do add-in em IndexedDB.</li>
              <li>Sem pasta e sem IndexedDB disponível: memória apenas como fallback técnico temporário.</li>
              <li>Desativado: sem storage intermedio e sem promessa de retoma local.</li>
            </ul>
          </InfoBlock>
          <ToggleRow
            label="Criar caso automaticamente ao abrir email novo"
            hint="Quando desligado, a aba deixa de gravar checkpoint intermédio novo em background ate haver save explicito."
            checked={draft.autoCreateCaseOnNewEmail}
            onChange={(next) => applyDraftPatch({ autoCreateCaseOnNewEmail: next })}
          />
          <ToggleRow
            label="Reabrir caso existente"
            hint="Controla a reidratação automática do caso local por `caseId`/`anchorEmailKey`."
            checked={draft.reopenExistingCase}
            onChange={(next) => applyDraftPatch({ reopenExistingCase: next })}
          />
          <ToggleRow
            label="Recriar copia intermedia a partir do servidor"
            hint="Controla se o histórico remoto volta a ser projetado para o intermédio quando ainda não existe cópia local."
            checked={draft.recreateIntermediateCopy}
            onChange={(next) => applyDraftPatch({ recreateIntermediateCopy: next })}
          />
          <ToggleRow
            label="Validar localizacao ao abrir"
            hint="Executa validação real da pasta intermédia ou do storage local do add-in ao abrir a aba."
            checked={draft.validateLocationOnOpen}
            onChange={(next) => applyDraftPatch({ validateLocationOnOpen: next })}
          />
          <ToggleRow
            label="Bloquear a aba se a localizacao falhar"
            hint="Quando ligado, uma falha de validação do intermédio bloqueia o fluxo Preparar."
            checked={draft.blockTabIfUnavailable}
            onChange={(next) => applyDraftPatch({ blockTabIfUnavailable: next })}
          />
          <ToggleRow
            label="Avisar indisponibilidade"
            hint="Emite aviso funcional quando a localização intermédia ativa fica indisponível."
            checked={draft.warnIfUnavailable}
            onChange={(next) => applyDraftPatch({ warnIfUnavailable: next })}
          />
          <ToggleRow
            label="Revalidar automaticamente"
            hint="Se a primeira validação falhar, tenta uma revalidação leve adicional."
            checked={draft.autoRetryValidation}
            onChange={(next) => applyDraftPatch({ autoRetryValidation: next })}
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
            <select
              style={S.select}
              value={draft.attachmentStrategy}
              onChange={(event) => applyDraftPatch({ attachmentStrategy: event.target.value as GroupsSettingsAttachmentStrategy })}
            >
              {Object.entries(ATTACHMENT_STRATEGY_LABELS).map(([value, label]) => (
                <option key={value} value={value}>
                  {label}
                </option>
              ))}
            </select>
          </FieldRow>
          <FieldRow label="Limite server (MB)" hint="Limiar hoje usado para decisoes best-effort no pipeline atual.">
            <input
              style={S.input}
              type="number"
              min={1}
              max={2048}
              value={draft.attachmentServerLimitMb}
              onChange={(event) => applyDraftPatch({ attachmentServerLimitMb: Number(event.target.value || 0) || 1 })}
            />
          </FieldRow>
          <FieldRow label="Limite intermadio (MB)" hint="Referencia guardada para a shell; nao representa um provider novo.">
            <input
              style={S.input}
              type="number"
              min={1}
              max={4096}
              value={draft.attachmentIntermediateLimitMb}
              onChange={(event) => applyDraftPatch({ attachmentIntermediateLimitMb: Number(event.target.value || 0) || 1 })}
            />
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
          <ToggleRow
            label="Guardar anexos no servidor"
            hint="Quando desligado, a persistencia final evita empurrar binario para o alvo server-first."
            checked={draft.saveAttachmentsOnServer}
            onChange={(next) => applyDraftPatch({ saveAttachmentsOnServer: next })}
          />
          <ToggleRow
            label="Guardar anexos fora do servidor"
            hint="Ativa persistencia file-backed quando existir path real acessivel ao runtime aprovado."
            checked={draft.saveAttachmentsOutsideServer}
            onChange={(next) => applyDraftPatch({ saveAttachmentsOutsideServer: next })}
          />
          <PathFieldRow
            label="Destino externo herdado"
            hint="Path real usado quando a politica de anexos escolhe o alvo externo."
            value={draft.externalAttachmentFolder}
            chooseLabel="Limpar"
            placeholder="ex.: C:/dados/grupos/anexos"
            onChangeText={(nextValue) => applyDraftPatch({ externalAttachmentFolder: nextValue })}
            onChoose={() => applyDraftPatch({ externalAttachmentFolder: "" })}
          />
          <ToggleRow
            label="Mostrar metadata no servidor"
            hint="Quando desligado, anexos puramente externos deixam de ser anunciados por metadata no payload final."
            checked={draft.showAttachmentMetadataOnServer}
            onChange={(next) => applyDraftPatch({ showAttachmentMetadataOnServer: next })}
          />
          <ToggleRow
            label="Preview imediato"
            hint="Força anexos com conteúdo a manterem preview imediato mesmo acima do limiar intermédio."
            checked={draft.requireImmediatePreview}
            onChange={(next) => applyDraftPatch({ requireImmediatePreview: next })}
          />
        </SectionShell>
      );
    }

    if (activeSection === "cleanup") {
      return (
        <SectionShell title="Limpeza" subtitle="Retencao local, cadencia de limpeza e protecao de casos mistos.">
          <FieldRow label="Dias para aviso de caso misto">
            <input style={S.input} type="number" min={1} max={3650} value={draft.mixedCaseWarningDays} onChange={(event) => applyDraftPatch({ mixedCaseWarningDays: Number(event.target.value || 0) || 1 })} />
          </FieldRow>
          <FieldRow label="Dias para aviso de abandono local">
            <input style={S.input} type="number" min={1} max={3650} value={draft.localAbandonedWarningDays} onChange={(event) => applyDraftPatch({ localAbandonedWarningDays: Number(event.target.value || 0) || 1 })} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso fechado">
            <input style={S.input} type="number" min={1} max={3650} value={draft.cleanupClosedCaseDays} onChange={(event) => applyDraftPatch({ cleanupClosedCaseDays: Number(event.target.value || 0) || 1 })} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de abandono">
            <input style={S.input} type="number" min={1} max={3650} value={draft.cleanupAbandonedCaseDays} onChange={(event) => applyDraftPatch({ cleanupAbandonedCaseDays: Number(event.target.value || 0) || 1 })} />
          </FieldRow>
          <FieldRow label="Frequencia">
            <select
              style={S.select}
              value={draft.cleanupFrequency}
              onChange={(event) => applyDraftPatch({ cleanupFrequency: event.target.value as GroupsSettingsFrequency })}
            >
              {Object.entries(FREQUENCY_LABELS).map(([value, label]) => (
                <option key={value} value={value}>{label}</option>
              ))}
            </select>
          </FieldRow>
          <ToggleRow
            label="Nunca apagar casos mistos em silencio"
            checked={draft.neverDeleteMixedSilently}
            onChange={(next) => applyDraftPatch({ neverDeleteMixedSilently: next })}
          />
          <div style={S.pathActions}>
            <ActionButton
              label={isSaving ? "A executar" : "Limpar agora"}
              disabled={isSaving || draft.storageMode === "disabled"}
              tone="danger"
              onClick={() => void handleCleanupNow()}
            />
          </div>
        </SectionShell>
      );
    }

    if (activeSection === "warnings") {
      return (
        <SectionShell title="Avisos" subtitle="Avisos leves locais controlados por frequencia e idade do caso.">
          <ToggleRow label="Avisar emails por classificar" checked={draft.warnUnclassifiedEmails} onChange={(next) => applyDraftPatch({ warnUnclassifiedEmails: next })} />
          <ToggleRow label="Avisar casos mistos" checked={draft.warnMixedCases} onChange={(next) => applyDraftPatch({ warnMixedCases: next })} />
          <FieldRow label="Frequencia">
            <select
              style={S.select}
              value={draft.warningFrequency}
              onChange={(event) => applyDraftPatch({ warningFrequency: event.target.value as GroupsSettingsFrequency })}
            >
              {Object.entries(FREQUENCY_LABELS).map(([value, label]) => (
                <option key={value} value={value}>{label}</option>
              ))}
            </select>
          </FieldRow>
          <ToggleRow label="Preparar ponte futura para tarefas" checked={draft.prepareTasksBridge} onChange={(next) => applyDraftPatch({ prepareTasksBridge: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "migration") {
      return (
        <SectionShell title="Migracao" subtitle="Migracao executavel do intermédio entre localizações reais, sem fingir migracao total do storage final.">
          <InfoBlock tone="warning" title="Migracao real do intermédio">
            <div style={S.inlineValue}>
              Nesta ronda a migracao real cobre o `IntermediateCase` entre a pasta intermédia definida e o fallback local do add-in. Migracao historica do storage final continua fora desta shell.
            </div>
          </InfoBlock>
          <PathFieldRow
            label="Destino guardado"
            hint="Pasta alvo para copiar/mover os casos intermédios persistidos."
            value={draft.migrationTarget}
            chooseLabel="Limpar"
            placeholder="ex.: C:/dados/grupos/intermedio-migrado"
            onChangeText={(nextValue) => applyDraftPatch({ migrationTarget: nextValue })}
            onChoose={() => applyDraftPatch({ migrationTarget: "" })}
          />
          <FieldRow label="Modo guardado">
            <select
              style={S.select}
              value={draft.migrationMode}
              onChange={(event) => applyDraftPatch({ migrationMode: event.target.value as GroupsSettingsMigrationMode })}
            >
              {Object.entries(MIGRATION_MODE_LABELS).map(([value, label]) => (
                <option key={value} value={value} disabled={value === "always_ask"}>
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
          <ToggleRow
            label="Seguranca estrita"
            hint="Bloqueia a migracao quando o destino ja contem cases com o mesmo identificador."
            checked={draft.strictMigrationSafety}
            onChange={(next) => applyDraftPatch({ strictMigrationSafety: next })}
          />
          <ToggleRow
            label="Fundir dados existentes"
            checked={draft.mergeExistingData}
            onChange={(next) => applyDraftPatch({ mergeExistingData: next })}
          />
          <div style={S.pathActions}>
            <ActionButton
              label={isSaving ? "A migrar" : "Migrar agora"}
              disabled={isSaving || !draft.migrationTarget || draft.migrationMode === "always_ask"}
              onClick={() => void handleMigrationNow()}
            />
          </div>
        </SectionShell>
      );
    }

    if (activeSection === "maintenance") {
      return (
        <SectionShell title="Manutencao" subtitle="Ferramentas locais da localização intermédia atual, sem sair do perimetro de Groups.">
          <ActionRow label="Revalidar storage intermedio" actionLabel={isSaving ? "A validar" : "Executar"} disabled={isSaving} onClick={() => void handleRevalidateStorage()} />
          <ActionRow label="Reconstruir indices locais" actionLabel={isSaving ? "A reconstruir" : "Executar"} disabled={isSaving || draft.storageMode === "disabled"} onClick={() => void handleRebuildIndexes()} />
          <ActionRow label="Limpar dados orfaos" actionLabel={isSaving ? "A limpar" : "Executar"} tone="danger" disabled={isSaving || draft.storageMode === "disabled"} onClick={() => void handleCleanupOrphans()} />
        </SectionShell>
      );
    }

    if (activeSection === "explore") {
      return (
        <SectionShell title="Explorar" subtitle="Bridges internas que já influenciam o bootstrap e as ações do studio de Groups.">
          <ToggleRow label="Servidor como base principal" checked={draft.explorerServerPrimary} onChange={(next) => applyDraftPatch({ explorerServerPrimary: next })} />
          <ToggleRow label="Abrir anexos guardados" checked={draft.explorerOpenStoredAttachments} onChange={(next) => applyDraftPatch({ explorerOpenStoredAttachments: next })} />
          <ToggleRow label="Permitir gerar resposta e reenvio" checked={draft.explorerGenerateReply} onChange={(next) => applyDraftPatch({ explorerGenerateReply: next })} />
        </SectionShell>
      );
    }

    return (
      <SectionShell title="Sobre" subtitle="Resumo da politica executavel de gravacao da aba Groups nesta fase.">
        <InfoBlock title="Intermedio">
          <ul style={S.infoList}>
            <li>`IntermediateCase` grava logo na pasta local definida quando ela existe.</li>
            <li>Sem pasta definida, o fallback principal é o storage local do add-in em IndexedDB.</li>
            <li>Memória só entra como fallback técnico quando o restante falha.</li>
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
  }, [activeSection, draft, handleCleanupNow, handleCleanupOrphans, handleMigrationNow, handleRebuildIndexes, handleRevalidateStorage, isSaving]);

  if (!open) return null;

  return (
    <div style={S.backdrop} role="presentation" onClick={() => !isSaving && onClose()}>
      <div style={S.modal} role="dialog" aria-modal="true" onClick={(event) => event.stopPropagation()}>
        <div style={S.modalHeader}>
          <div style={S.modalHeaderText}>
            <div style={S.modalEyebrow}>Groups</div>
            <div style={S.modalTitle}>Settings da aba Groups</div>
          </div>
          <div style={S.modalActions}>
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
        <div style={S.modalBody}>
          <div style={S.sidebar}>
            {SECTION_ENTRIES.map((entry) => (
              <SectionButton key={entry.id} entry={entry} active={entry.id === activeSection} onClick={() => setActiveSection(entry.id)} />
            ))}
          </div>
          <div style={S.content}>{content}</div>
        </div>
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
  modalHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 10,
    padding: "10px 12px",
    borderBottom: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.82)",
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
  sidebar: {
    display: "grid",
    alignContent: "start",
    gap: 4,
    padding: 10,
    borderRight: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.72)",
    overflowY: "auto",
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
    gridTemplateColumns: "minmax(0, 1fr) minmax(180px, 240px)",
    gap: 10,
    alignItems: "center",
    padding: "7px 8px",
    borderRadius: 12,
    border: "1px solid rgba(148,163,184,0.14)",
    background: "rgba(255,255,255,0.88)",
  },
  rowDisabled: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) minmax(180px, 240px)",
    gap: 10,
    alignItems: "center",
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
    justifyContent: "flex-end",
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
