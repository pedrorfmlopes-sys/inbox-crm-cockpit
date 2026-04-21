import React, { useEffect, useMemo, useState } from "react";
import type { CockpitSettingsV1, GroupStorageMode } from "@/settings";
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
  storageMode: GroupStorageMode;
  baseFolderPath: string;
  locationStatus: string;
  autoCreateCaseOnNewEmail: boolean;
  reopenExistingCase: boolean;
  recreateIntermediateCopy: boolean;
  validateLocationOnOpen: boolean;
  blockTabIfUnavailable: boolean;
  warnIfUnavailable: boolean;
  autoRetryValidation: boolean;
  attachmentStrategy: "metadata_first" | "hybrid_guarded" | "external_first";
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
  migrationMode: "confirm" | "move" | "copy";
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
  return {
    storageMode: storage?.mode || "supabase",
    baseFolderPath: storage?.baseFolderPath || storage?.chosenFolder?.path || storage?.localDevice?.rootPath || "",
    locationStatus: storage?.baseFolderPath ? "Localização pronta para validar" : "Localização ainda por confirmar",
    autoCreateCaseOnNewEmail: true,
    reopenExistingCase: true,
    recreateIntermediateCopy: false,
    validateLocationOnOpen: true,
    blockTabIfUnavailable: false,
    warnIfUnavailable: true,
    autoRetryValidation: true,
    attachmentStrategy: "metadata_first",
    saveAttachmentsOnServer: Boolean(storage?.supabase?.promoteAttachmentMetadataOnSave),
    saveAttachmentsOutsideServer: true,
    attachmentServerLimitMb: Number(storage?.attachmentPromptThresholdMb || 10),
    attachmentIntermediateLimitMb: 25,
    externalAttachmentFolder: storage?.chosenFolder?.path || storage?.baseFolderPath || "",
    showAttachmentMetadataOnServer: true,
    requireImmediatePreview: false,
    mixedCaseWarningDays: 10,
    localAbandonedWarningDays: 14,
    cleanupClosedCaseDays: 45,
    cleanupAbandonedCaseDays: 30,
    cleanupFrequency: "weekly",
    neverDeleteMixedSilently: true,
    warnUnclassifiedEmails: true,
    warnMixedCases: true,
    warningFrequency: "daily",
    prepareTasksBridge: false,
    migrationTarget: storage?.chosenFolder?.path || storage?.baseFolderPath || "",
    migrationMode: "confirm",
    allowMoveExistingData: false,
    strictMigrationSafety: true,
    mergeExistingData: false,
    explorerServerPrimary: storage?.mode === "supabase" || storage?.mode === "hybrid",
    explorerOpenStoredAttachments: true,
    explorerGenerateReply: false,
    groupsVersion: "Grupos v1 / shell de settings",
    quickDiagnostic: storage?.mode ? `Modo ativo detectado: ${storage.mode}` : "Sem modo ativo resolvido",
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
      <button type="button" style={tone === "danger" ? S.actionButtonDanger : S.actionButton}>
        {actionLabel}
      </button>
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
        <SectionShell title="General" subtitle="Entrada compacta da aba Grupos e estado base desta configuração.">
          <FieldRow label="Aba Grupos ativa" hint="Indicador visual da área de settings que estamos a preparar para a aba Groups.">
            <div style={S.inlineValue}>Mostrar settings da aba Grupos</div>
          </FieldRow>
          <FieldRow label="Estado da shell" hint="Esta ronda cria apenas a estrutura visual e os placeholders seguros para futuras ligações.">
            <div style={S.statusPill}>UI pronta para ligação futura</div>
          </FieldRow>
        </SectionShell>
      );
    }

    if (activeSection === "intermediate_storage") {
      return (
        <SectionShell title="Armazenamento intermédio" subtitle="Shell visual da base intermédia, sem ligar ainda migração ou validação real.">
          <FieldRow label="Modo de armazenamento" hint="Escolha visual do modo ativo para a base intermédia de Grupos.">
            <select style={S.select} value={draft.storageMode} onChange={(event) => setDraft((current) => ({ ...current, storageMode: event.target.value as GroupStorageMode }))}>
              <option value="supabase">Tudo no Supabase</option>
              <option value="local_device">Local neste PC</option>
              <option value="chosen_folder">Local em pasta escolhida</option>
              <option value="hybrid">Híbrido</option>
            </select>
          </FieldRow>
          <FieldRow label="Pasta base" hint="Caminho base mostrado de forma compacta; a ligação real entra noutra ronda.">
            <input style={S.input} value={draft.baseFolderPath} onChange={(event) => setDraft((current) => ({ ...current, baseFolderPath: event.target.value }))} placeholder="C:\\Base\\Grupos" />
          </FieldRow>
          <div style={S.actionGrid}>
            <ActionRow label="Escolher localização" hint="Abre o picker real numa fase posterior." actionLabel="Escolher" />
            <ActionRow label="Validar" hint="Validação real da localização fica preparada para uma ronda própria." actionLabel="Validar" />
            <ActionRow label="Abrir pasta" hint="Abertura real da localização ainda não está ligada nesta shell." actionLabel="Abrir pasta" />
          </div>
          <FieldRow label="Estado da ligação" hint="Diagnóstico rápido do target intermédio sem executar validação destrutiva.">
            <div style={S.inlineValue}>{draft.locationStatus}</div>
          </FieldRow>
          <ToggleRow label="Criar caso automaticamente ao abrir email novo" hint="Mantido como placeholder controlado; sem automatismo real nesta ronda." checked={draft.autoCreateCaseOnNewEmail} onChange={(next) => setDraft((current) => ({ ...current, autoCreateCaseOnNewEmail: next }))} />
          <ToggleRow label="Reabrir caso existente a partir da base intermédia" hint="Prepara a intenção de retoma sem abrir já a lógica de matching final." checked={draft.reopenExistingCase} onChange={(next) => setDraft((current) => ({ ...current, reopenExistingCase: next }))} />
          <ToggleRow label="Recriar cópia intermédia quando o histórico só existir no servidor" hint="Só estrutura visual nesta fase." checked={draft.recreateIntermediateCopy} onChange={(next) => setDraft((current) => ({ ...current, recreateIntermediateCopy: next }))} />
          <ToggleRow label="Validar a localização ao abrir a aba Grupos" hint="O toggle fica pronto; a validação real não entra nesta ronda." checked={draft.validateLocationOnOpen} onChange={(next) => setDraft((current) => ({ ...current, validateLocationOnOpen: next }))} />
          <ToggleRow label="Bloquear a aba Grupos se a localização não estiver acessível" hint="Guard rail visual para uma política futura." checked={draft.blockTabIfUnavailable} onChange={(next) => setDraft((current) => ({ ...current, blockTabIfUnavailable: next }))} />
          <ToggleRow label="Mostrar aviso se a pasta deixar de estar acessível" hint="Aviso subtil preparado para futura ligação." checked={draft.warnIfUnavailable} onChange={(next) => setDraft((current) => ({ ...current, warnIfUnavailable: next }))} />
          <ToggleRow label="Tentar revalidar automaticamente" hint="Comportamento ainda stubado." checked={draft.autoRetryValidation} onChange={(next) => setDraft((current) => ({ ...current, autoRetryValidation: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "attachments") {
      return (
        <SectionShell title="Anexos" subtitle="Shell da política de anexos, ainda sem ligar storage final nem regras destrutivas.">
          <FieldRow label="Estratégia de armazenamento" hint="Placeholder da política principal para anexos classificados.">
            <select style={S.select} value={draft.attachmentStrategy} onChange={(event) => setDraft((current) => ({ ...current, attachmentStrategy: event.target.value as GroupsSettingsDraft["attachmentStrategy"] }))}>
              <option value="metadata_first">Metadata first</option>
              <option value="hybrid_guarded">Híbrida controlada</option>
              <option value="external_first">Externa primeiro</option>
            </select>
          </FieldRow>
          <ToggleRow label="Guardar anexos no servidor" hint="Só toggle de interface nesta ronda; sem upload real." checked={draft.saveAttachmentsOnServer} onChange={(next) => setDraft((current) => ({ ...current, saveAttachmentsOnServer: next }))} />
          <ToggleRow label="Guardar anexos fora do servidor" hint="Mantém a shell pronta para o fluxo de pasta externa." checked={draft.saveAttachmentsOutsideServer} onChange={(next) => setDraft((current) => ({ ...current, saveAttachmentsOutsideServer: next }))} />
          <FieldRow label="Limite para guardar no servidor (MB)" hint="Valor visual para futura regra de promoção controlada.">
            <input style={S.input} type="number" min={1} value={draft.attachmentServerLimitMb} onChange={(event) => setDraft((current) => ({ ...current, attachmentServerLimitMb: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Limite intermédio opcional (MB)" hint="Faixa de decisão intermédia ainda sem lógica final.">
            <input style={S.input} type="number" min={0} value={draft.attachmentIntermediateLimitMb} onChange={(event) => setDraft((current) => ({ ...current, attachmentIntermediateLimitMb: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Pasta externa de anexos classificados" hint="Caminho placeholder para futura ligação com storage externo.">
            <input style={S.input} value={draft.externalAttachmentFolder} onChange={(event) => setDraft((current) => ({ ...current, externalAttachmentFolder: event.target.value }))} placeholder="C:\\Anexos\\Classificados" />
          </FieldRow>
          <ToggleRow label="Mostrar sempre metadados de todos os anexos no servidor" hint="Garante consistência visual do inventário, sem promotion automática nesta ronda." checked={draft.showAttachmentMetadataOnServer} onChange={(next) => setDraft((current) => ({ ...current, showAttachmentMetadataOnServer: next }))} />
          <ToggleRow label="Exigir preview imediato para anexos marcados como guardados" hint="Placeholder para a política futura de preview obrigatório." checked={draft.requireImmediatePreview} onChange={(next) => setDraft((current) => ({ ...current, requireImmediatePreview: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "cleanup") {
      return (
        <SectionShell title="Limpeza" subtitle="Parâmetros de retenção e limpeza preparados sem executar rotinas reais.">
          <FieldRow label="Dias para aviso de caso misto" hint="Lead time antes do aviso visual de casos mistos.">
            <input style={S.input} type="number" min={1} value={draft.mixedCaseWarningDays} onChange={(event) => setDraft((current) => ({ ...current, mixedCaseWarningDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para aviso de caso local abandonado" hint="Janela de warning antes de qualquer limpeza futura.">
            <input style={S.input} type="number" min={1} value={draft.localAbandonedWarningDays} onChange={(event) => setDraft((current) => ({ ...current, localAbandonedWarningDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso fechado" hint="Sem cron real nesta fase; apenas shell visual do setting.">
            <input style={S.input} type="number" min={1} value={draft.cleanupClosedCaseDays} onChange={(event) => setDraft((current) => ({ ...current, cleanupClosedCaseDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Dias para limpeza de caso local abandonado" hint="Threshold visual preparado para futura rotina segura.">
            <input style={S.input} type="number" min={1} value={draft.cleanupAbandonedCaseDays} onChange={(event) => setDraft((current) => ({ ...current, cleanupAbandonedCaseDays: Number(event.target.value || 0) }))} />
          </FieldRow>
          <FieldRow label="Frequência da verificação" hint="A rotina real fica para uma ronda própria.">
            <select style={S.select} value={draft.cleanupFrequency} onChange={(event) => setDraft((current) => ({ ...current, cleanupFrequency: event.target.value as GroupsSettingsDraft["cleanupFrequency"] }))}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Nunca apagar em silêncio casos mistos" hint="Botão de segurança preparado para ligação posterior." checked={draft.neverDeleteMixedSilently} onChange={(next) => setDraft((current) => ({ ...current, neverDeleteMixedSilently: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "warnings") {
      return (
        <SectionShell title="Avisos" subtitle="Preferências compactas para avisos de trabalho pendente ou casos mistos.">
          <ToggleRow label="Avisar emails por classificar" hint="A shell guarda a preferência visual, sem scheduler real nesta ronda." checked={draft.warnUnclassifiedEmails} onChange={(next) => setDraft((current) => ({ ...current, warnUnclassifiedEmails: next }))} />
          <ToggleRow label="Avisar casos mistos sem atividade" hint="Preparado para futuras rotinas de aviso." checked={draft.warnMixedCases} onChange={(next) => setDraft((current) => ({ ...current, warnMixedCases: next }))} />
          <FieldRow label="Frequência dos avisos" hint="Escolha de cadência ainda sem automação ligada.">
            <select style={S.select} value={draft.warningFrequency} onChange={(event) => setDraft((current) => ({ ...current, warningFrequency: event.target.value as GroupsSettingsDraft["warningFrequency"] }))}>
              <option value="manual">Manual</option>
              <option value="daily">Diária</option>
              <option value="weekly">Semanal</option>
            </select>
          </FieldRow>
          <ToggleRow label="Preparar integração com futura área de tarefas" hint="Marcador visual para a futura ponte com Tarefas, sem abrir essa frente já." checked={draft.prepareTasksBridge} onChange={(next) => setDraft((current) => ({ ...current, prepareTasksBridge: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "migration") {
      return (
        <SectionShell title="Migração" subtitle="Ações de migração preparadas visualmente, sem mover dados nesta ronda.">
          <FieldRow label="Alterar localização da base intermédia" hint="Novo destino alvo mostrado de forma simples para futura ligação.">
            <input style={S.input} value={draft.migrationTarget} onChange={(event) => setDraft((current) => ({ ...current, migrationTarget: event.target.value }))} placeholder="Nova localização base" />
          </FieldRow>
          <FieldRow label="Ao alterar localização" hint="Define a atitude base da migração sem executar nada agora.">
            <select style={S.select} value={draft.migrationMode} onChange={(event) => setDraft((current) => ({ ...current, migrationMode: event.target.value as GroupsSettingsDraft["migrationMode"] }))}>
              <option value="confirm">Confirmar antes de agir</option>
              <option value="move">Mover quando confirmado</option>
              <option value="copy">Criar cópia quando confirmado</option>
            </select>
          </FieldRow>
          <ToggleRow label="Permitir mover dados existentes" hint="Comportamento destrutivo mantido apenas como intenção visual." checked={draft.allowMoveExistingData} onChange={(next) => setDraft((current) => ({ ...current, allowMoveExistingData: next }))} />
          <ToggleRow label="Regra de segurança na migração" hint="Quando ativo, obriga confirmação forte e validação antes de qualquer ação real futura." checked={draft.strictMigrationSafety} onChange={(next) => setDraft((current) => ({ ...current, strictMigrationSafety: next }))} />
          <ToggleRow label="Fundir com dados já existentes na nova pasta" hint="Mantido como opção discreta, ainda sem merge real." checked={draft.mergeExistingData} onChange={(next) => setDraft((current) => ({ ...current, mergeExistingData: next }))} />
        </SectionShell>
      );
    }

    if (activeSection === "maintenance") {
      return (
        <SectionShell title="Manutenção" subtitle="Botões preparados visualmente e destacados com disciplina, sem ações destrutivas reais.">
          <ActionRow label="Criar backup" hint="A infraestrutura real de backup entra noutra ronda." actionLabel="Criar backup" />
          <ActionRow label="Repor backup" hint="Preparado para futura recuperação guiada." actionLabel="Repor backup" />
          <ActionRow label="Reset da base intermédia" hint="Ação destrutiva ainda sem ligação operacional." actionLabel="Reset base" tone="danger" />
          <ActionRow label="Reset do servidor" hint="Mantido apenas como shell visual; sem backend destrutivo ligado." actionLabel="Reset servidor" tone="danger" />
          <ActionRow label="Reset total" hint="Ação mais sensível, ainda só visual." actionLabel="Reset total" tone="danger" />
          <ActionRow label="Refazer categorização" hint="Atalho preparado para fluxo futuro." actionLabel="Refazer" />
          <ActionRow label="Revalidar dados" hint="Preparado para uma rotina futura de diagnóstico." actionLabel="Revalidar" />
        </SectionShell>
      );
    }

    if (activeSection === "explore") {
      return (
        <SectionShell title="Explorar" subtitle="Preferências visuais para a futura frente de Explorar, sem a abrir já.">
          <ToggleRow label="Usar servidor como base principal do Explorador" hint="Somente preferência visual nesta ronda." checked={draft.explorerServerPrimary} onChange={(next) => setDraft((current) => ({ ...current, explorerServerPrimary: next }))} />
          <ToggleRow label="Permitir abrir anexos guardados" hint="Abertura real de anexos fica noutra fase." checked={draft.explorerOpenStoredAttachments} onChange={(next) => setDraft((current) => ({ ...current, explorerOpenStoredAttachments: next }))} />
          <ToggleRow label="Permitir gerar resposta e reenvio" hint="Só placeholder visual; sem lógica de resposta nesta frente." checked={draft.explorerGenerateReply} onChange={(next) => setDraft((current) => ({ ...current, explorerGenerateReply: next }))} />
        </SectionShell>
      );
    }

    return (
      <SectionShell title="Sobre" subtitle="Resumo rápido desta shell de settings da aba Grupos.">
        <FieldRow label="Versão do módulo Grupos" hint="Identificador simples desta ronda de interface.">
          <div style={S.inlineValue}>{draft.groupsVersion}</div>
        </FieldRow>
        <FieldRow label="Diagnóstico rápido" hint="Texto curto para futura leitura de estado técnico do módulo.">
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
    fontSize: 13,
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
    fontSize: 9,
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
    fontSize: 11,
    fontWeight: 650,
    color: "#243244",
  },
  sectionSubtitle: {
    fontSize: 8.6,
    color: "#64748b",
    lineHeight: 1.35,
  },
  sectionBody: {
    display: "grid",
    gap: 7,
  },
  row: {
    display: "grid",
    gridTemplateColumns: "minmax(0, 1fr) minmax(160px, 220px)",
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
    fontSize: 8.8,
    fontWeight: 600,
    color: "#3a495c",
    lineHeight: 1.3,
  },
  rowControl: {
    minWidth: 0,
  },
  input: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.2,
    color: "#243244",
    boxSizing: "border-box",
  },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.24)",
    background: "#fff",
    padding: "6px 8px",
    fontSize: 9.2,
    color: "#243244",
    boxSizing: "border-box",
  },
  inlineValue: {
    fontSize: 9.1,
    color: "#334155",
    lineHeight: 1.35,
  },
  statusPill: {
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    padding: "4px 8px",
    borderRadius: 999,
    background: "rgba(219,234,254,0.68)",
    color: "#1d4ed8",
    fontSize: 8.4,
    fontWeight: 700,
  },
  toggleOn: {
    width: 28,
    height: 16,
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
    width: 28,
    height: 16,
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
    width: 12,
    height: 12,
    borderRadius: 999,
    background: "#fff",
    boxShadow: "0 1px 2px rgba(15,23,42,0.18)",
  },
  actionGrid: {
    display: "grid",
    gap: 7,
  },
  actionButton: {
    borderRadius: 10,
    border: "1px solid rgba(148,163,184,0.22)",
    background: "#fff",
    color: "#334155",
    padding: "5px 8px",
    fontSize: 8.8,
    fontWeight: 650,
    cursor: "pointer",
  },
  actionButtonDanger: {
    borderRadius: 10,
    border: "1px solid rgba(220,38,38,0.16)",
    background: "rgba(254,242,242,0.94)",
    color: "#b91c1c",
    padding: "5px 8px",
    fontSize: 8.8,
    fontWeight: 650,
    cursor: "pointer",
  },
};

export default GroupsSettingsPanel;
