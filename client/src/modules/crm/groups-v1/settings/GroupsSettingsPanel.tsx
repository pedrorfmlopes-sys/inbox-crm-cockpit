import React, { useEffect, useMemo, useState } from "react";
import { HelpHint } from "@/ui/HelpHint";
import * as Icons from "@/ui/icons";
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

function buildDraft(value: GroupsTabSettings | null | undefined): GroupsTabSettings {
  return normalizeGroupsTabSettings(value || null);
}

function promptForPath(label: string, currentValue: string): string | null {
  if (typeof window === "undefined" || typeof window.prompt !== "function") return null;
  const nextValue = window.prompt(`Defina ${label.toLowerCase()}`, currentValue || "");
  if (nextValue == null) return null;
  return nextValue.trim();
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
  onClick,
}: {
  label: string;
  tone?: "neutral" | "danger";
  onClick?: () => void;
}) {
  return (
    <button type="button" style={tone === "danger" ? S.actionButtonDanger : S.actionButton} onClick={onClick}>
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
  onChoose,
}: {
  label: string;
  hint?: string;
  value: string;
  chooseLabel?: string;
  showValidate?: boolean;
  showOpen?: boolean;
  onChoose?: () => void;
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
          <ActionButton label={chooseLabel} onClick={onChoose} />
          {showValidate ? <ActionButton label="Validar" /> : null}
          {showOpen ? <ActionButton label="Abrir pasta" /> : null}
        </div>
      </div>
    </div>
  );
}

export function GroupsSettingsPanel({ open, value, onClose, onSave }: Props): JSX.Element | null {
  const [activeSection, setActiveSection] = useState<GroupsSettingsSection>("general");
  const [draft, setDraft] = useState<GroupsTabSettings>(() => buildDraft(value));
  const [isSaving, setIsSaving] = useState(false);

  const applyDraftPatch = (patch: Partial<GroupsTabSettings>) => {
    setDraft((current) => normalizeGroupsTabSettings({ ...current, ...patch }));
  };

  useEffect(() => {
    if (!open) return;
    setDraft(buildDraft(value));
    setActiveSection("general");
    setIsSaving(false);
  }, [open, value]);

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

  const content = useMemo(() => {
    if (activeSection === "general") {
      return (
        <SectionShell title="General" subtitle="Estado base desta configuração da aba Grupos.">
          <ToggleRow
            label="Aba Grupos ativa"
            hint="Mantém esta aba ativa nas definições da aplicação."
            checked={draft.groupsTabEnabled}
            onChange={(next) => applyDraftPatch({ groupsTabEnabled: next })}
          />
          <FieldRow label="Estado" hint="Resumo curto desta configuração.">
            <div style={S.inlineValue}>{draft.groupsTabEnabled ? "Aba ativa nesta configuração" : "Aba desativada nesta configuração"}</div>
          </FieldRow>
        </SectionShell>
      );
    }

    if (activeSection === "intermediate_storage") {
      return (
        <SectionShell title="Armazenamento intermédio" subtitle="Preferências persistidas da base intermédia, sem abrir ligação real nesta ronda.">
          <FieldRow label="Modo de armazenamento" hint="Escolhe se a base intermédia fica ativa para OneDrive / SharePoint ou desativada.">
            <select
              style={S.select}
              value={draft.storageMode}
              onChange={(event) => applyDraftPatch({ storageMode: event.target.value as GroupsSettingsStorageMode })}
            >
              <option value="onedrive_sharepoint">OneDrive / SharePoint</option>
              <option value="disabled">Desativado</option>
            </select>
          </FieldRow>
          <PathFieldRow
            label="Pasta base"
            hint="Mostra a localização base preparada para a aba Grupos."
            value={draft.baseFolderPath}
            onChoose={() => {
              const nextPath = promptForPath("a localização da pasta base", draft.baseFolderPath);
              if (nextPath == null) return;
              applyDraftPatch({ baseFolderPath: nextPath });
            }}
          />
          <FieldRow label="Estado da ligação" hint="Resumo leve da localização configurada, sem validação real da pasta.">
            <div style={S.inlineValue}>{draft.locationStatus}</div>
          </FieldRow>
          <ToggleRow label="Criar caso automaticamente ao abrir email novo" hint="Prepara a abertura direta de um novo caso para a aba Grupos." checked={draft.autoCreateCaseOnNewEmail} onChange={(next) => applyDraftPatch({ autoCreateCaseOnNewEmail: next })} />
          <ToggleRow label="Reabrir caso existente a partir da base intermédia" hint="Mantém ativa a retoma de um caso já presente na base intermédia." checked={draft.reopenExistingCase} onChange={(next) => applyDraftPatch({ reopenExistingCase: next })} />
          <ToggleRow label="Recriar cópia intermédia quando o histórico só existir no servidor" hint="Mantém preparada uma cópia local de trabalho quando o histórico só existir no servidor." checked={draft.recreateIntermediateCopy} onChange={(next) => applyDraftPatch({ recreateIntermediateCopy: next })} />
          <ToggleRow label="Validar a localização ao abrir a aba Grupos" hint="Guarda a preferência de validação logo ao abrir esta aba." checked={draft.validateLocationOnOpen} onChange={(next) => applyDraftPatch({ validateLocationOnOpen: next })} />
          <ToggleRow label="Bloquear a aba Grupos se a localização não estiver acessível" hint="Mantém a preferência de bloqueio quando a localização principal não responder." checked={draft.blockTabIfUnavailable} onChange={(next) => applyDraftPatch({ blockTabIfUnavailable: next })} />
          <ToggleRow label="Mostrar aviso se a pasta deixar de estar acessível" hint="Mostra um aviso curto se a localização configurada deixar de responder." checked={draft.warnIfUnavailable} onChange={(next) => applyDraftPatch({ warnIfUnavailable: next })} />
          <ToggleRow label="Tentar revalidar automaticamente" hint="Guarda a preferência de nova tentativa automática após uma falha." checked={draft.autoRetryValidation} onChange={(next) => applyDraftPatch({ autoRetryValidation: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "attachments") {
      return (
        <SectionShell title="Anexos" subtitle="Preferências persistidas da política de anexos, ainda sem storage final ligado.">
          <FieldRow label="Estratégia de armazenamento" hint="Define a regra principal que a app deverá aplicar aos anexos classificados.">
            <select
              style={S.select}
              value={draft.attachmentStrategy}
              onChange={(event) => applyDraftPatch({ attachmentStrategy: event.target.value as GroupsSettingsAttachmentStrategy })}
            >
              <option value="server">Todos no servidor</option>
              <option value="outside">Todos fora do servidor</option>
              <option value="by_size">Por tamanho</option>
            </select>
          </FieldRow>
          <ToggleRow label="Guardar anexos no servidor" hint="Mantém a preferência principal para anexos guardados no servidor." checked={draft.saveAttachmentsOnServer} onChange={(next) => applyDraftPatch({ saveAttachmentsOnServer: next })} />
          <ToggleRow label="Guardar anexos fora do servidor" hint="Mantém a preferência para anexos guardados fora do servidor." checked={draft.saveAttachmentsOutsideServer} onChange={(next) => applyDraftPatch({ saveAttachmentsOutsideServer: next })} />
          <FieldRow label="Limite para guardar no servidor (MB)" hint="Valor usado como referência para a regra por tamanho.">
            <input style={S.input} type="number" min={1} value={draft.attachmentServerLimitMb} onChange={(event) => applyDraftPatch({ attachmentServerLimitMb: Number(event.target.value || 0) })} />
          </FieldRow>
          <FieldRow label="Limite intermédio opcional (MB)" hint="Faixa intermédia preparada para a regra por tamanho.">
            <input style={S.input} type="number" min={1} value={draft.attachmentIntermediateLimitMb} onChange={(event) => applyDraftPatch({ attachmentIntermediateLimitMb: Number(event.target.value || 0) })} />
          </FieldRow>
          <PathFieldRow
            label="Pasta externa de anexos classificados"
            hint="Mostra a localização prevista para anexos guardados fora do servidor."
            value={draft.externalAttachmentFolder}
            onChoose={() => {
              const nextPath = promptForPath("a pasta externa de anexos classificados", draft.externalAttachmentFolder);
              if (nextPath == null) return;
              applyDraftPatch({ externalAttachmentFolder: nextPath });
            }}
          />
          <ToggleRow label="Mostrar sempre metadados de todos os anexos no servidor" hint="Mantém visível o inventário de anexos mesmo quando o ficheiro ficar fora do servidor." checked={draft.showAttachmentMetadataOnServer} onChange={(next) => applyDraftPatch({ showAttachmentMetadataOnServer: next })} />
          <ToggleRow label="Exigir preview imediato para anexos marcados como guardados" hint="Guarda a preferência para preview imediato dos anexos assinalados." checked={draft.requireImmediatePreview} onChange={(next) => applyDraftPatch({ requireImmediatePreview: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "cleanup") {
      return (
        <SectionShell title="Limpeza" subtitle="Parâmetros persistidos de aviso e limpeza, ainda sem rotinas reais.">
          <FieldRow label="Dias para aviso de caso misto" hint="Número de dias até aparecer o aviso de caso misto.">
            <input style={S.input} type="number" min={1} value={draft.mixedCaseWarningDays} onChange={(event) => applyDraftPatch({ mixedCaseWarningDays: Number(event.target.value || 0) })} />
          </FieldRow>
          <FieldRow label="Dias para aviso de caso local abandonado" hint="Número de dias até aparecer o aviso para um caso local sem atividade.">
            <input style={S.input} type="number" min={1} value={draft.localAbandonedWarningDays} onChange={(event) => applyDraftPatch({ localAbandonedWarningDays: Number(event.target.value || 0) })} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso fechado" hint="Prazo previsto antes da limpeza de um caso fechado.">
            <input style={S.input} type="number" min={1} value={draft.cleanupClosedCaseDays} onChange={(event) => applyDraftPatch({ cleanupClosedCaseDays: Number(event.target.value || 0) })} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso local abandonado" hint="Prazo previsto antes da limpeza de um caso local abandonado.">
            <input style={S.input} type="number" min={1} value={draft.cleanupAbandonedCaseDays} onChange={(event) => applyDraftPatch({ cleanupAbandonedCaseDays: Number(event.target.value || 0) })} />
          </FieldRow>
          <FieldRow label="Frequência da verificação" hint="Cadência prevista para esta verificação.">
            <select style={S.select} value={draft.cleanupFrequency} onChange={(event) => applyDraftPatch({ cleanupFrequency: event.target.value as GroupsSettingsFrequency })}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Nunca apagar em silêncio casos mistos" hint="Mantém confirmação visível antes de qualquer limpeza deste tipo." checked={draft.neverDeleteMixedSilently} onChange={(next) => applyDraftPatch({ neverDeleteMixedSilently: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "warnings") {
      return (
        <SectionShell title="Avisos" subtitle="Preferências persistidas de aviso, sem disparo real nesta ronda.">
          <ToggleRow label="Avisar emails por classificar" hint="Guarda a preferência de aviso para emails ainda por classificar." checked={draft.warnUnclassifiedEmails} onChange={(next) => applyDraftPatch({ warnUnclassifiedEmails: next })} />
          <ToggleRow label="Avisar casos mistos sem atividade" hint="Guarda a preferência de aviso para casos mistos sem atividade recente." checked={draft.warnMixedCases} onChange={(next) => applyDraftPatch({ warnMixedCases: next })} />
          <FieldRow label="Frequência dos avisos" hint="Cadência prevista para estes avisos.">
            <select style={S.select} value={draft.warningFrequency} onChange={(event) => applyDraftPatch({ warningFrequency: event.target.value as GroupsSettingsFrequency })}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Preparar integração com futura área de tarefas" hint="Reserva esta preferência para a futura área de tarefas." checked={draft.prepareTasksBridge} onChange={(next) => applyDraftPatch({ prepareTasksBridge: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "migration") {
      return (
        <SectionShell title="Migração" subtitle="Preferências persistidas de migração, sem mover dados nesta ronda.">
          <PathFieldRow
            label="Alterar localização da base intermédia"
            hint="Mostra o destino previsto para a base intermédia."
            value={draft.migrationTarget}
            onChoose={() => {
              const nextPath = promptForPath("a nova localização da base intermédia", draft.migrationTarget);
              if (nextPath == null) return;
              applyDraftPatch({ migrationTarget: nextPath });
            }}
          />
          <FieldRow label="Ao alterar localização" hint="Define como a app deve reagir antes de qualquer migração real.">
            <select style={S.select} value={draft.migrationMode} onChange={(event) => applyDraftPatch({ migrationMode: event.target.value as GroupsSettingsMigrationMode })}>
              <option value="always_ask">Perguntar sempre</option>
              <option value="move">Mover quando confirmado</option>
              <option value="copy">Criar cópia quando confirmado</option>
            </select>
          </FieldRow>
          <ToggleRow label="Permitir mover dados existentes" hint="Guarda a preferência para mover dados atuais quando a migração for confirmada." checked={draft.allowMoveExistingData} onChange={(next) => applyDraftPatch({ allowMoveExistingData: next })} />
          <FieldRow label="Regra de segurança na migração" hint="Mantém esta regra ativa e só de leitura nesta shell.">
            <div style={S.inlineValue}>{draft.strictMigrationSafety ? "Ativa (só leitura)" : "Desativada"}</div>
          </FieldRow>
          <ToggleRow label="Fundir com dados já existentes na nova pasta" hint="Decide se a nova localização pode aproveitar dados já existentes." checked={draft.mergeExistingData} onChange={(next) => applyDraftPatch({ mergeExistingData: next })} />
        </SectionShell>
      );
    }

    if (activeSection === "maintenance") {
      return (
        <SectionShell title="Manutenção" subtitle="Ações preparadas visualmente, sem execução real nesta ronda.">
          <ActionRow label="Criar backup" hint="Prepara a criação de um backup da base intermédia." actionLabel="Criar backup" />
          <ActionRow label="Repor backup" hint="Prepara a reposição de um backup existente." actionLabel="Repor backup" />
          <ActionRow label="Reset da base intermédia" hint="Ação sensível, mantida apenas como botão visual." actionLabel="Reset base" tone="danger" />
          <ActionRow label="Reset do servidor" hint="Ação sensível, sem ligação real nesta shell." actionLabel="Reset servidor" tone="danger" />
          <ActionRow label="Reset total" hint="Ação mais sensível, mantida apenas como placeholder visual." actionLabel="Reset total" tone="danger" />
          <ActionRow label="Refazer categorização" hint="Prepara uma revalidação de categorias numa ronda própria." actionLabel="Refazer" />
          <ActionRow label="Revalidar dados" hint="Prepara um diagnóstico curto dos dados atuais." actionLabel="Revalidar" />
        </SectionShell>
      );
    }

    if (activeSection === "explore") {
      return (
        <SectionShell title="Explorar" subtitle="Preferências persistidas da futura frente de Explorar, sem a abrir já.">
          <ToggleRow label="Usar servidor como base principal do Explorador" hint="Guarda a preferência principal de base para o Explorador." checked={draft.explorerServerPrimary} onChange={(next) => applyDraftPatch({ explorerServerPrimary: next })} />
          <ToggleRow label="Permitir abrir anexos guardados" hint="Mantém ativa a possibilidade de abrir anexos já guardados." checked={draft.explorerOpenStoredAttachments} onChange={(next) => applyDraftPatch({ explorerOpenStoredAttachments: next })} />
          <ToggleRow label="Permitir gerar resposta e reenvio" hint="Reserva a preferência para resposta e reenvio futuros." checked={draft.explorerGenerateReply} onChange={(next) => applyDraftPatch({ explorerGenerateReply: next })} />
        </SectionShell>
      );
    }

    return (
      <SectionShell title="Sobre" subtitle="Resumo curto desta janela de configuração da aba Grupos.">
        <FieldRow label="Versão do módulo Grupos" hint="Identificador simples desta configuração.">
          <div style={S.inlineValue}>{draft.groupsVersion}</div>
        </FieldRow>
        <FieldRow label="Diagnóstico rápido" hint="Resumo leve do estado atual desta área.">
          <div style={S.inlineValue}>{draft.quickDiagnostic}</div>
        </FieldRow>
      </SectionShell>
    );
  }, [activeSection, draft]);

  if (!open) return null;

  return (
    <div style={S.overlay} onClick={() => !isSaving && onClose()}>
      <div style={S.modal} onClick={(event) => event.stopPropagation()}>
        <div style={S.modalHeader}>
          <div style={S.modalTitleWrap}>
            <div style={S.modalKicker}>Groups</div>
            <div style={S.modalTitle}>Settings</div>
          </div>
          <div style={S.modalActions}>
            <button type="button" style={S.headerButtonSecondary} onClick={onClose} disabled={isSaving}>Fechar</button>
            <button type="button" style={isSaving ? S.headerButtonPrimaryDisabled : S.headerButtonPrimary} onClick={() => void handleSave()} disabled={isSaving}>
              <Icons.Save size={12} />
              {isSaving ? "A guardar..." : "Guardar"}
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
