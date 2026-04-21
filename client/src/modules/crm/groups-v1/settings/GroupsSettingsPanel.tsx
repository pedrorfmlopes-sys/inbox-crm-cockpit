import React, { useEffect, useMemo, useState } from "react";
import type { CockpitSettingsV1 } from "@/settings";
import { HelpHint } from "@/ui/HelpHint";
import * as Icons from "@/ui/icons";

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

type GroupsSettingsDraft = {
  groupsTabEnabled: boolean;
  storageMode: "onedrive_sharepoint" | "disabled";
  baseFolderPath: string;
  locationStatus: string;
  autoCreateCaseOnNewEmail: boolean;
  reopenExistingCase: boolean;
  recreateIntermediateCopy: boolean;
  validateLocationOnOpen: boolean;
  blockTabIfUnavailable: boolean;
  warnIfUnavailable: boolean;
  autoRetryValidation: boolean;
  attachmentStrategy: "server" | "outside" | "by_size";
  saveAttachmentsOnServer: boolean;
  saveAttachmentsOutsideServer: boolean;
  attachmentServerLimitMb: number;
  attachmentIntermediateLimitMb: number;
  externalAttachmentFolder: string;
  showAttachmentMetadataOnServer: boolean;
  requireImmediatePreview: boolean;
  mixedCaseWarningDays: number;
  localAbandonedWarningDays: number;
  cleanupClosedCaseDays: number;
  cleanupAbandonedCaseDays: number;
  cleanupFrequency: "manual" | "daily" | "weekly";
  neverDeleteMixedSilently: boolean;
  warnUnclassifiedEmails: boolean;
  warnMixedCases: boolean;
  warningFrequency: "manual" | "daily" | "weekly";
  prepareTasksBridge: boolean;
  migrationTarget: string;
  migrationMode: "always_ask" | "move" | "copy";
  allowMoveExistingData: boolean;
  strictMigrationSafety: boolean;
  mergeExistingData: boolean;
  explorerServerPrimary: boolean;
  explorerOpenStoredAttachments: boolean;
  explorerGenerateReply: boolean;
  groupsVersion: string;
  quickDiagnostic: string;
};

type Props = {
  open: boolean;
  settings: CockpitSettingsV1 | null;
  onClose: () => void;
  onSave: () => void;
};

type SectionEntry = {
  id: GroupsSettingsSection;
  label: string;
  icon: React.ReactNode;
};

const SECTION_ENTRIES: SectionEntry[] = [
  { id: "general", label: "General", icon: <Icons.Settings size={12} /> },
  { id: "intermediate_storage", label: "Armazenamento intermédio", icon: <Icons.Database size={12} /> },
  { id: "attachments", label: "Anexos", icon: <Icons.Paperclip size={12} /> },
  { id: "cleanup", label: "Limpeza", icon: <Icons.RefreshCw size={12} /> },
  { id: "warnings", label: "Avisos", icon: <Icons.AlertCircle size={12} /> },
  { id: "migration", label: "Migração", icon: <Icons.Upload size={12} /> },
  { id: "maintenance", label: "Manutenção", icon: <Icons.Trash size={12} /> },
  { id: "explore", label: "Explorar", icon: <Icons.Search size={12} /> },
  { id: "about", label: "Sobre", icon: <Icons.MessageSquare size={12} /> },
];

function buildDraft(settings: CockpitSettingsV1 | null): GroupsSettingsDraft {
  const storage = settings?.groupStorage;
  const storageMode = storage?.mode === "chosen_folder" || storage?.mode === "hybrid"
    ? "onedrive_sharepoint"
    : "disabled";

  return {
    groupsTabEnabled: true,
    storageMode,
    baseFolderPath: storage?.baseFolderPath || storage?.chosenFolder?.path || storage?.localDevice?.rootPath || "",
    locationStatus: storage?.baseFolderPath ? "Localização configurada para validação." : "Localização ainda por definir.",
    autoCreateCaseOnNewEmail: true,
    reopenExistingCase: true,
    recreateIntermediateCopy: true,
    validateLocationOnOpen: true,
    blockTabIfUnavailable: true,
    warnIfUnavailable: true,
    autoRetryValidation: true,
    attachmentStrategy: "by_size",
    saveAttachmentsOnServer: true,
    saveAttachmentsOutsideServer: true,
    attachmentServerLimitMb: Number(storage?.attachmentPromptThresholdMb || 10),
    attachmentIntermediateLimitMb: 50,
    externalAttachmentFolder: storage?.chosenFolder?.path || storage?.baseFolderPath || "",
    showAttachmentMetadataOnServer: true,
    requireImmediatePreview: true,
    mixedCaseWarningDays: 15,
    localAbandonedWarningDays: 30,
    cleanupClosedCaseDays: 15,
    cleanupAbandonedCaseDays: 90,
    cleanupFrequency: "daily",
    neverDeleteMixedSilently: true,
    warnUnclassifiedEmails: true,
    warnMixedCases: true,
    warningFrequency: "daily",
    prepareTasksBridge: false,
    migrationTarget: storage?.chosenFolder?.path || storage?.baseFolderPath || "",
    migrationMode: "always_ask",
    allowMoveExistingData: false,
    strictMigrationSafety: true,
    mergeExistingData: false,
    explorerServerPrimary: true,
    explorerOpenStoredAttachments: true,
    explorerGenerateReply: true,
    groupsVersion: "Grupos v1",
    quickDiagnostic: storage?.baseFolderPath ? "Base intermédia pronta para validação." : "Base intermédia ainda sem localização definida.",
  };
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

function FieldRow({
  label,
  hint,
  children,
}: {
  label: string;
  hint?: string;
  children: React.ReactNode;
}) {
  return (
    <label style={S.row}>
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
  onChange,
}: {
  label: string;
  hint?: string;
  checked: boolean;
  onChange: (next: boolean) => void;
}) {
  return (
    <div style={S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <button type="button" style={checked ? S.toggleOn : S.toggleOff} onClick={() => onChange(!checked)} aria-pressed={checked}>
        <span style={S.toggleThumb} />
      </button>
    </div>
  );
}

function ActionButton({
  label,
  tone = "neutral",
}: {
  label: string;
  tone?: "neutral" | "danger";
}) {
  return (
    <button type="button" style={tone === "danger" ? S.actionButtonDanger : S.actionButton}>
      {label}
    </button>
  );
}

function ActionRow({
  label,
  hint,
  actionLabel,
  tone = "neutral",
}: {
  label: string;
  hint?: string;
  actionLabel: string;
  tone?: "neutral" | "danger";
}) {
  return (
    <div style={S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <ActionButton label={actionLabel} tone={tone} />
    </div>
  );
}

function PathFieldRow({
  label,
  hint,
  value,
  chooseLabel = "Escolher localização",
  showValidate = true,
  showOpen = true,
}: {
  label: string;
  hint?: string;
  value: string;
  chooseLabel?: string;
  showValidate?: boolean;
  showOpen?: boolean;
}) {
  return (
    <div style={S.row}>
      <div style={S.rowLabelWrap}>
        <span style={S.rowLabel}>{label}</span>
        {hint ? <SettingHint text={hint} /> : null}
      </div>
      <div style={S.pathControl}>
        <div style={S.pathValue} title={value || "Sem localização definida"}>
          {value || "Sem localização definida"}
        </div>
        <div style={S.pathActions}>
          <ActionButton label={chooseLabel} />
          {showValidate ? <ActionButton label="Validar" /> : null}
          {showOpen ? <ActionButton label="Abrir pasta" /> : null}
        </div>
      </div>
    </div>
  );
}

export function GroupsSettingsPanel({ open, settings, onClose, onSave }: Props): JSX.Element | null {
  const [activeSection, setActiveSection] = useState<GroupsSettingsSection>("general");
  const [draft, setDraft] = useState<GroupsSettingsDraft>(() => buildDraft(settings));

  useEffect(() => {
    if (!open) return;
    setDraft(buildDraft(settings));
    setActiveSection("general");
  }, [open, settings]);

  useEffect(() => {
    if (!open) return;
    const handleKeyDown = (event: KeyboardEvent) => {
      if (event.key === "Escape") onClose();
    };
    window.addEventListener("keydown", handleKeyDown);
    return () => window.removeEventListener("keydown", handleKeyDown);
  }, [open, onClose]);

  const content = useMemo(() => {
    if (activeSection === "general") {
      return (
        <SectionShell title="General" subtitle="Estado base desta janela de configuração da aba Grupos.">
          <ToggleRow
            label="Aba Grupos ativa"
            hint="Controla visualmente se a aba Grupos está disponível nesta configuração."
            checked={draft.groupsTabEnabled}
            onChange={(next) => setDraft((current) => ({ ...current, groupsTabEnabled: next }))}
          />
          <FieldRow label="Estado" hint="Resumo curto do estado atual desta configuração.">
            <div style={S.inlineValue}>Configuração disponível nesta janela.</div>
          </FieldRow>
        </SectionShell>
      );
    }

    if (activeSection === "intermediate_storage") {
      return (
        <SectionShell title="Armazenamento intermédio" subtitle="Preferências visuais da base intermédia, sem ligar ainda validação ou migração real.">
          <FieldRow label="Modo de armazenamento" hint="Escolhe se a base intermédia usa OneDrive/SharePoint ou fica desativada.">
            <select
              style={S.select}
              value={draft.storageMode}
              onChange={(event) =>
                setDraft((current) => ({ ...current, storageMode: event.target.value as GroupsSettingsDraft["storageMode"] }))
              }
            >
              <option value="onedrive_sharepoint">OneDrive / SharePoint</option>
              <option value="disabled">Desativado</option>
            </select>
          </FieldRow>
          <PathFieldRow
            label="Pasta base"
            hint="Mostra a localização base que será usada por esta configuração."
            value={draft.baseFolderPath}
          />
          <FieldRow label="Estado da ligação" hint="Mostra um resumo curto da localização configurada.">
            <div style={S.inlineValue}>{draft.locationStatus}</div>
          </FieldRow>
          <ToggleRow label="Criar caso automaticamente ao abrir email novo" hint="Quando ligado, prepara a abertura direta de um novo caso." checked={draft.autoCreateCaseOnNewEmail} onChange={(next) => setDraft((current) => ({ ...current, autoCreateCaseOnNewEmail: next }))} />
          <ToggleRow label="Reabrir caso existente a partir da base intermédia" hint="Quando ligado, prepara a retoma de um caso já encontrado na base intermédia." checked={draft.reopenExistingCase} onChange={(next) => setDraft((current) => ({ ...current, reopenExistingCase: next }))} />
          <ToggleRow label="Recriar cópia intermédia quando o histórico só existir no servidor" hint="Mantém uma cópia de trabalho quando só existir histórico remoto." checked={draft.recreateIntermediateCopy} onChange={(next) => setDraft((current) => ({ ...current, recreateIntermediateCopy: next }))} />
          <ToggleRow label="Validar a localização ao abrir a aba Grupos" hint="Confirma a localização assim que a aba for aberta." checked={draft.validateLocationOnOpen} onChange={(next) => setDraft((current) => ({ ...current, validateLocationOnOpen: next }))} />
          <ToggleRow label="Bloquear a aba Grupos se a localização não estiver acessível" hint="Evita continuar se a localização principal não responder." checked={draft.blockTabIfUnavailable} onChange={(next) => setDraft((current) => ({ ...current, blockTabIfUnavailable: next }))} />
          <ToggleRow label="Mostrar aviso se a pasta deixar de estar acessível" hint="Mostra um aviso curto quando a localização falhar." checked={draft.warnIfUnavailable} onChange={(next) => setDraft((current) => ({ ...current, warnIfUnavailable: next }))} />
          <ToggleRow label="Tentar revalidar automaticamente" hint="Tenta confirmar outra vez a localização quando houver falha." checked={draft.autoRetryValidation} onChange={(next) => setDraft((current) => ({ ...current, autoRetryValidation: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "attachments") {
      return (
        <SectionShell title="Anexos" subtitle="Preferências visuais da política de anexos, sem ligar ainda storage final.">
          <FieldRow label="Estratégia de armazenamento" hint="Define a regra principal de destino dos anexos classificados.">
            <select
              style={S.select}
              value={draft.attachmentStrategy}
              onChange={(event) =>
                setDraft((current) => ({ ...current, attachmentStrategy: event.target.value as GroupsSettingsDraft["attachmentStrategy"] }))
              }
            >
              <option value="server">Todos no servidor</option>
              <option value="outside">Todos fora do servidor</option>
              <option value="by_size">Por tamanho</option>
            </select>
          </FieldRow>
          <ToggleRow label="Guardar anexos no servidor" hint="Mostra a preferência principal para anexos no servidor." checked={draft.saveAttachmentsOnServer} onChange={(next) => setDraft((current) => ({ ...current, saveAttachmentsOnServer: next }))} />
          <ToggleRow label="Guardar anexos fora do servidor" hint="Mostra a preferência para anexos guardados fora do servidor." checked={draft.saveAttachmentsOutsideServer} onChange={(next) => setDraft((current) => ({ ...current, saveAttachmentsOutsideServer: next }))} />
          <FieldRow label="Limite para guardar no servidor (MB)" hint="Acima deste valor, a decisão pode seguir outra regra.">
            <input style={S.input} type="number" min={1} value={draft.attachmentServerLimitMb} onChange={(event) => setDraft((current) => ({ ...current, attachmentServerLimitMb: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Limite intermédio opcional (MB)" hint="Faixa intermédia para a regra por tamanho.">
            <input style={S.input} type="number" min={0} value={draft.attachmentIntermediateLimitMb} onChange={(event) => setDraft((current) => ({ ...current, attachmentIntermediateLimitMb: Number(event.target.value || 0) }))} />
          </FieldRow>
          <PathFieldRow
            label="Pasta externa de anexos classificados"
            hint="Mostra o destino preparado para anexos guardados fora do servidor."
            value={draft.externalAttachmentFolder}
          />
          <ToggleRow label="Mostrar sempre metadados de todos os anexos no servidor" hint="Mantém visível o inventário dos anexos, mesmo quando o ficheiro ficar fora do servidor." checked={draft.showAttachmentMetadataOnServer} onChange={(next) => setDraft((current) => ({ ...current, showAttachmentMetadataOnServer: next }))} />
          <ToggleRow label="Exigir preview imediato para anexos marcados como guardados" hint="Sinaliza que o preview deve ficar disponível logo que o anexo for marcado." checked={draft.requireImmediatePreview} onChange={(next) => setDraft((current) => ({ ...current, requireImmediatePreview: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "cleanup") {
      return (
        <SectionShell title="Limpeza" subtitle="Parâmetros de aviso e limpeza, ainda sem executar rotinas reais.">
          <FieldRow label="Dias para aviso de caso misto" hint="Dias até aparecer o aviso de caso misto.">
            <input style={S.input} type="number" min={1} value={draft.mixedCaseWarningDays} onChange={(event) => setDraft((current) => ({ ...current, mixedCaseWarningDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para aviso de caso local abandonado" hint="Dias até aparecer o aviso para um caso local sem atividade.">
            <input style={S.input} type="number" min={1} value={draft.localAbandonedWarningDays} onChange={(event) => setDraft((current) => ({ ...current, localAbandonedWarningDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso fechado" hint="Prazo previsto antes da limpeza de um caso fechado.">
            <input style={S.input} type="number" min={1} value={draft.cleanupClosedCaseDays} onChange={(event) => setDraft((current) => ({ ...current, cleanupClosedCaseDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso local abandonado" hint="Prazo previsto antes da limpeza de um caso local abandonado.">
            <input style={S.input} type="number" min={1} value={draft.cleanupAbandonedCaseDays} onChange={(event) => setDraft((current) => ({ ...current, cleanupAbandonedCaseDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Frequência da verificação" hint="Cadência prevista para esta verificação.">
            <select style={S.select} value={draft.cleanupFrequency} onChange={(event) => setDraft((current) => ({ ...current, cleanupFrequency: event.target.value as GroupsSettingsDraft["cleanupFrequency"] }))}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Nunca apagar em silêncio casos mistos" hint="Mantém confirmação visível antes de qualquer limpeza deste tipo." checked={draft.neverDeleteMixedSilently} onChange={(next) => setDraft((current) => ({ ...current, neverDeleteMixedSilently: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "warnings") {
      return (
        <SectionShell title="Avisos" subtitle="Preferências visuais para avisos de trabalho pendente ou casos mistos.">
          <ToggleRow label="Avisar emails por classificar" hint="Mostra a preferência de aviso para emails ainda por classificar." checked={draft.warnUnclassifiedEmails} onChange={(next) => setDraft((current) => ({ ...current, warnUnclassifiedEmails: next }))} />
          <ToggleRow label="Avisar casos mistos sem atividade" hint="Mostra a preferência de aviso para casos mistos sem atividade recente." checked={draft.warnMixedCases} onChange={(next) => setDraft((current) => ({ ...current, warnMixedCases: next }))} />
          <FieldRow label="Frequência dos avisos" hint="Cadência prevista para estes avisos.">
            <select style={S.select} value={draft.warningFrequency} onChange={(event) => setDraft((current) => ({ ...current, warningFrequency: event.target.value as GroupsSettingsDraft["warningFrequency"] }))}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Preparar integração com futura área de tarefas" hint="Reserva esta preferência para a futura área de tarefas." checked={draft.prepareTasksBridge} onChange={(next) => setDraft((current) => ({ ...current, prepareTasksBridge: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "migration") {
      return (
        <SectionShell title="Migração" subtitle="Ações de migração preparadas visualmente, sem mover dados nesta ronda.">
          <PathFieldRow
            label="Alterar localização da base intermédia"
            hint="Mostra o destino previsto para a base intermédia."
            value={draft.migrationTarget}
          />
          <FieldRow label="Ao alterar localização" hint="Define como a app deve reagir antes de qualquer migração.">
            <select style={S.select} value={draft.migrationMode} onChange={(event) => setDraft((current) => ({ ...current, migrationMode: event.target.value as GroupsSettingsDraft["migrationMode"] }))}>
              <option value="always_ask">Perguntar sempre</option>
              <option value="move">Mover quando confirmado</option>
              <option value="copy">Criar cópia quando confirmado</option>
            </select>
          </FieldRow>
          <ToggleRow label="Permitir mover dados existentes" hint="Mantém a possibilidade de mover dados atuais quando a migração for confirmada." checked={draft.allowMoveExistingData} onChange={(next) => setDraft((current) => ({ ...current, allowMoveExistingData: next }))} />
          <FieldRow label="Regra de segurança na migração" hint="Mantém esta regra ativa e só de leitura nesta shell.">
            <div style={S.inlineValue}>Ativa (só leitura)</div>
          </FieldRow>
          <ToggleRow label="Fundir com dados já existentes na nova pasta" hint="Decide se a nova localização pode aproveitar dados já existentes." checked={draft.mergeExistingData} onChange={(next) => setDraft((current) => ({ ...current, mergeExistingData: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "maintenance") {
      return (
        <SectionShell title="Manutenção" subtitle="Ações preparadas visualmente, sem executar operações reais nesta ronda.">
          <ActionRow label="Criar backup" hint="Prepara a criação de um backup da base intermédia." actionLabel="Criar backup" />
          <ActionRow label="Repor backup" hint="Prepara a reposição de um backup existente." actionLabel="Repor backup" />
          <ActionRow label="Reset da base intermédia" hint="Ação sensível, mantida apenas como botão visual." actionLabel="Reset base" tone="danger" />
          <ActionRow label="Reset do servidor" hint="Ação sensível, sem ligação real nesta shell." actionLabel="Reset servidor" tone="danger" />
          <ActionRow label="Reset total" hint="Ação mais sensível, mantida apenas como placeholder visual." actionLabel="Reset total" tone="danger" />
          <ActionRow label="Refazer categorização" hint="Prepara uma revalidação de categorias no futuro." actionLabel="Refazer" />
          <ActionRow label="Revalidar dados" hint="Prepara um diagnóstico curto dos dados atuais." actionLabel="Revalidar" />
        </SectionShell>
      );
    }

    if (activeSection === "explore") {
      return (
        <SectionShell title="Explorar" subtitle="Preferências visuais para a futura frente de Explorar, sem a abrir já.">
          <ToggleRow label="Usar servidor como base principal do Explorador" hint="Mostra a preferência principal de base para o Explorador." checked={draft.explorerServerPrimary} onChange={(next) => setDraft((current) => ({ ...current, explorerServerPrimary: next }))} />
          <ToggleRow label="Permitir abrir anexos guardados" hint="Mantém aberta a possibilidade de abrir anexos já guardados." checked={draft.explorerOpenStoredAttachments} onChange={(next) => setDraft((current) => ({ ...current, explorerOpenStoredAttachments: next }))} />
          <ToggleRow label="Permitir gerar resposta e reenvio" hint="Reserva a preferência para resposta e reenvio futuros." checked={draft.explorerGenerateReply} onChange={(next) => setDraft((current) => ({ ...current, explorerGenerateReply: next }))} />
        </SectionShell>
      );
    }

    return (
      <SectionShell title="Sobre" subtitle="Resumo curto desta janela de configuração da aba Grupos.">
        <FieldRow label="Versão do módulo Grupos" hint="Identificador simples desta configuração.">
          <div style={S.inlineValue}>{draft.groupsVersion}</div>
        </FieldRow>
        <FieldRow label="Diagnóstico rápido" hint="Resumo curto do estado atual desta área.">
          <div style={S.inlineValue}>{draft.quickDiagnostic}</div>
        </FieldRow>
      </SectionShell>
    );
  }, [activeSection, draft]);

  if (!open) return null;

  return (
    <div style={S.overlay} onClick={onClose}>
      <div style={S.modal} onClick={(event) => event.stopPropagation()}>
        <div style={S.modalHeader}>
          <div style={S.modalTitleWrap}>
            <div style={S.modalKicker}>Groups</div>
            <div style={S.modalTitle}>Settings</div>
          </div>
          <div style={S.modalActions}>
            <button type="button" style={S.headerButtonSecondary} onClick={onClose}>Fechar</button>
            <button type="button" style={S.headerButtonPrimary} onClick={onSave}>
              <Icons.Save size={12} />
              Guardar
            </button>
          </div>
        </div>

        <div style={S.modalBody}>
          <aside style={S.sidebar}>
            {SECTION_ENTRIES.map((entry) => (
              <SectionButton key={entry.id} entry={entry} active={entry.id === activeSection} onClick={() => setActiveSection(entry.id)} />
            ))}
          </aside>
          <main style={S.content}>{content}</main>
        </div>
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  overlay: {
    position: "fixed",
    inset: 0,
    zIndex: 60,
    background: "rgba(15,23,42,0.18)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    padding: 10,
    boxSizing: "border-box",
  },
  modal: {
    width: "min(720px, 100%)",
    maxHeight: "min(680px, calc(100vh - 20px))",
    display: "grid",
    gridTemplateRows: "auto minmax(0, 1fr)",
    borderRadius: 18,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "rgba(248,250,252,0.98)",
    boxShadow: "0 18px 40px rgba(15,23,42,0.16)",
    overflow: "hidden",
  },
  modalHeader: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 10,
    padding: "10px 12px",
    borderBottom: "1px solid rgba(148,163,184,0.16)",
    background: "rgba(255,255,255,0.9)",
  },
  modalTitleWrap: {
    display: "grid",
    gap: 1,
    minWidth: 0,
  },
  modalKicker: {
    fontSize: 8,
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
    fontSize: 9.4,
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
    fontSize: 8.9,
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
  rowLabelWrap: {
    display: "inline-flex",
    alignItems: "center",
    gap: 5,
    minWidth: 0,
  },
  rowLabel: {
    fontSize: 9.2,
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
    fontSize: 9.5,
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
    fontSize: 9.5,
    color: "#243244",
    boxSizing: "border-box",
  },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.5,
    color: "#243244",
    boxSizing: "border-box",
  },
  inlineValue: {
    fontSize: 9.4,
    color: "#334155",
    lineHeight: 1.35,
  },
  toggleOn: {
    width: 30,
    height: 17,
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
    width: 30,
    height: 17,
    borderRadius: 999,
    border: "1px solid rgba(239,68,68,0.18)",
    background: "rgba(239,68,68,0.58)",
    padding: 1,
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "flex-start",
    cursor: "pointer",
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
};

export default GroupsSettingsPanel;
