import React, { useEffect, useMemo, useState } from "react";
import {
  clearSignatureImageDataUrl,
  getSettings,
  getSignatureImageDataUrl,
  resetSettings,
  saveSettings,
  setSignatureImageDataUrl,
  type AppLocale,
  type CockpitSettingsV1,
  type Crm2OdooLayoutTarget,
  type GroupStorageMode,
  type LangOption,
  type ReferenceEntityKey,
  type ReplyLength,
  type SkinId,
} from "../settings";
import { applySkin } from "./skins";
import * as Icons from "./icons";
import { useCockpit } from "../components/shell/CockpitProvider";
import { aiListModels, validateCrm2OdooLayout, type Crm2LayoutValidationResult } from "../api";
import { GROUP_STORAGE_MODE_LABELS } from "../modules/crm/groups-v1/storage/modes";
import { resolveGroupStorageRuntime } from "../modules/crm/groups-v1/storage/resolveStorageMode";
import { buildGroupStorageValidationPayload, describeGroupStorageCapabilities } from "../modules/crm/groups-v1/storage/storageCapabilities";
import { validateGroupStorageTarget, type GroupStorageValidationResult } from "../modules/crm/groups-v1/storage/worksetApi";
import { PanelState, type PanelStateTone } from "./PanelState";
import { previewReferenceCode } from "../referenceCodes";

type StatusNotice = { tone: PanelStateTone; title: string; description?: string };
type StatusValue = StatusNotice | string | null;

const LOCALE_LABEL: Record<AppLocale, string> = {
  "pt-PT": "Português (Portugal)",
  "es-ES": "Espanhol (Espanha)",
  "en-GB": "Inglês (UK)",
  "it-IT": "Italiano (IT)",
  "de-DE": "Alemão (DE)",
};

const LANG_OPTIONS: Array<{ value: LangOption; label: string }> = [
  { value: "auto", label: "Auto" },
  { value: "pt-PT", label: "Português (PT)" },
  { value: "es-ES", label: "Espanhol (ES)" },
  { value: "en-GB", label: "Inglês (UK)" },
  { value: "it-IT", label: "Italiano (IT)" },
  { value: "de-DE", label: "Alemão (DE)" },
];

const PICKER_LANGS: AppLocale[] = ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"];

const LENGTH_OPTIONS: Array<{ value: ReplyLength; label: string }> = [
  { value: "xs", label: "Extra curta" },
  { value: "s", label: "Curta" },
  { value: "m", label: "Média" },
  { value: "l", label: "Longa" },
];

const TONE_OPTIONS = [
  { value: "neutro", label: "Neutro" },
  { value: "curto", label: "Curto" },
  { value: "direto", label: "Direto" },
  { value: "simpático", label: "Simpático" },
] as const;

const SKIN_OPTIONS: Array<{ value: SkinId; label: string }> = [
  { value: "classic", label: "Classic" },
  { value: "mailmaestro", label: "MailMaestro" },
  { value: "vibrant", label: "Vibrant (Cockpit 3.0)" },
];

const EXECUTABLE_GROUP_STORAGE_MODE_OPTIONS: Array<{ value: GroupStorageMode; label: string }> = [
  { value: "supabase", label: "Cockpit Cloud" },
  { value: "local_device", label: "Local acessivel ao servidor" },
  { value: "chosen_folder", label: "Pasta local / sincronizada" },
  { value: "hybrid", label: "Hibrido" },
];

const GROUP_STORAGE_WEB_URL_BLOCKER_FACTS = [
  "O manifest do add-in no repo so declara ReadWriteMailbox.",
  "O runtime Graph atual em office.ts so pede Mail.Read, User.Read e People.Read.",
  "O backend atual grava binario por filesystem; nao existe uploader Graph/SharePoint por URL web.",
];

const REFERENCE_ENTITY_LABELS: Record<ReferenceEntityKey, string> = {
  lead: "Lead",
  project: "Projeto",
  task: "Tarefa",
  ticket: "Ticket",
};

function localeShort(loc: AppLocale): string {
  if (loc === "pt-PT") return "PT";
  if (loc === "es-ES") return "ES";
  if (loc === "en-GB") return "EN";
  if (loc === "it-IT") return "IT";
  if (loc === "de-DE") return "DE";
  return loc;
}

export function SettingsPanel(): JSX.Element {
  const { settingsSection: section, setSettingsSection: setSection } = useCockpit();
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<StatusValue>(null);
  const [model, setModel] = useState<CockpitSettingsV1 | null>(null);

  // local-only uploaded signature images (dataURL), per locale
  const [sigImgLocal, setSigImgLocal] = useState<Partial<Record<AppLocale, string>>>({});
  const [availableModels, setAvailableModels] = useState<{ openai: string[]; gemini: string[] }>({ openai: [], gemini: [] });
  const [fetchingModels, setFetchingModels] = useState(false);
  const [groupStorageValidation, setGroupStorageValidation] = useState<GroupStorageValidationResult | null>(null);
  const [validatingGroupStorage, setValidatingGroupStorage] = useState(false);

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const s = await getSettings();
        if (!alive) return;
        setModel(s);

        // Load local (dataURL) signature images
        const map: Partial<Record<AppLocale, string>> = {};
        for (const loc of PICKER_LANGS) map[loc] = getSignatureImageDataUrl(loc) || "";
        setSigImgLocal(map);

        try {
          if (s.skinId) applySkin(s.skinId);
        } catch {
          /* ignore */
        }
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  const fetchModels = async () => {
    setFetchingModels(true);
    try {
      const res = await aiListModels();
      if (res.ok) {
        setAvailableModels({ openai: res.openai, gemini: res.gemini });
      }
    } catch (e) {
      console.warn("Failed to fetch models:", e);
    } finally {
      setFetchingModels(false);
    }
  };

  useEffect(() => {
    if (section === "conns") {
      fetchModels();
    }
  }, [section]);

  const title = useMemo(() => {
    if (section === "general") return "Geral";
    if (section === "conns") return "Ligações";
    if (section === "ai") return "IA Knowledge";
    if (section === "persona") return "Minha Persona";
    if (section === "signature") return "Assinatura";
    if (section === "references") return "Codigos de Referencia";
    if (section === "groups") return "Grupos";
    if (section === "crm2layout") return "CRM2 / Odoo Layout";
    return "Proteção (O Moat)";
  }, [section]);

  const resolvedGroupStorage = useMemo(
    () => (model ? resolveGroupStorageRuntime(model) : null),
    [model]
  );
  const groupStoragePathRequired = model ? model.groupStorage.mode !== "supabase" : false;
  const groupStorageCapabilities = useMemo(
    () => (model && resolvedGroupStorage
      ? describeGroupStorageCapabilities({
          settings: model.groupStorage,
          runtime: resolvedGroupStorage,
          validation: groupStorageValidation,
        })
      : []),
    [groupStorageValidation, model, resolvedGroupStorage]
  );
  const groupStorageRequiresGraphOrAdmin = resolvedGroupStorage?.projectSupport.requiresGraphOrAdmin === true;

  const validateCurrentGroupStorage = async (settingsModel: CockpitSettingsV1): Promise<GroupStorageValidationResult | null> => {
    if (settingsModel.groupStorage.mode === "supabase") {
      const result: GroupStorageValidationResult = {
        mode: "supabase",
        provider: "cloud",
        fileBacked: false,
        supported: true,
        basePath: "",
        normalizedBasePath: "",
        isWebUrl: false,
        requiresServerAccessiblePath: false,
        canStoreManifest: true,
        canStoreBinary: true,
        pickerAvailable: false,
        notes: ["Modo cloud validado localmente."],
      };
      setGroupStorageValidation(result);
      return result;
    }

    setValidatingGroupStorage(true);
    try {
      const result = await validateGroupStorageTarget(buildGroupStorageValidationPayload(settingsModel.groupStorage));
      setGroupStorageValidation(result);
      return result;
    } finally {
      setValidatingGroupStorage(false);
    }
  };

  useEffect(() => {
    if (!model || section !== "groups") return;
    void validateCurrentGroupStorage(model);
  }, [
    model,
    model?.groupStorage.mode,
    model?.groupStorage.baseFolderPath,
    model?.groupStorage.chosenFolder.kind,
    model?.groupStorage.hybrid.primaryTarget,
    section,
  ]);

  async function onSave() {
    if (!model) return;
    setSaving(true);
    setStatus(null);
    try {
      let nextModel = model;
      const validation = await validateCurrentGroupStorage(model);
      if (model.groupStorage.mode !== "supabase") {
        if (!validation?.supported) {
          setStatus({
            tone: "error",
            title: "Destino de storage bloqueado",
            description: validation?.blockingReason || "Define um destino realmente acessivel ao servidor antes de guardar.",
          });
          return;
        }
        const normalizedBasePath = String(validation.normalizedBasePath || model.groupStorage.baseFolderPath || "").trim();
        nextModel = {
          ...model,
          groupStorage: {
            ...model.groupStorage,
            baseFolderPath: normalizedBasePath,
            localDevice: {
              ...model.groupStorage.localDevice,
              rootPath:
                model.groupStorage.mode === "local_device" || model.groupStorage.hybrid.primaryTarget === "local_device"
                  ? normalizedBasePath
                  : model.groupStorage.localDevice.rootPath,
            },
            chosenFolder: {
              ...model.groupStorage.chosenFolder,
              path:
                model.groupStorage.mode === "chosen_folder" || model.groupStorage.hybrid.primaryTarget === "chosen_folder"
                  ? normalizedBasePath
                  : model.groupStorage.chosenFolder.path,
            },
          },
        };
      }
      await saveSettings(nextModel);
      setModel(nextModel);
      setStatus({ tone: "success", title: "Definições guardadas", description: "As alterações já estão disponíveis no cockpit." });
      setTimeout(() => setStatus(null), 1800);
    } catch (e: any) {
      setStatus({ tone: "error", title: "Falha ao guardar", description: e?.message || "Não foi possível guardar as definições." });
    } finally {
      setSaving(false);
    }
  }

  async function onReset() {
    setSaving(true);
    setStatus(null);
    try {
      const s = await resetSettings();
      setModel(s);

      // reset does not remove local-only images automatically (by design)
      // keep current local preview synced
      const map: Partial<Record<AppLocale, string>> = {};
      for (const loc of PICKER_LANGS) map[loc] = getSignatureImageDataUrl(loc) || "";
      setSigImgLocal(map);

      setStatus({ tone: "success", title: "Definições repostas", description: "Os valores guardados voltaram ao estado por defeito." });
      setTimeout(() => setStatus(null), 2200);
    } catch (e: any) {
      setStatus({ tone: "error", title: "Falha ao repor", description: e?.message || "Não foi possível repor as definições." });
    } finally {
      setSaving(false);
    }
  }

  function setSigUrl(loc: AppLocale, url: string) {
    if (!model) return;
    setModel({
      ...model,
      signatureImageUrl: { ...(model.signatureImageUrl || {}), [loc]: url },
    });
  }

  function setSigMaxW(loc: AppLocale, w: number) {
    if (!model) return;
    const safe = Math.max(120, Math.min(900, Number.isFinite(w) ? w : 260));
    setModel({
      ...model,
      signatureImageMaxWidth: { ...(model.signatureImageMaxWidth || {}), [loc]: safe },
    });
  }

  function onUploadSig(loc: AppLocale, file: File) {
    const reader = new FileReader();
    reader.onload = () => {
      const dataUrl = String(reader.result || "").trim();
      if (!dataUrl) return;
      setSignatureImageDataUrl(loc, dataUrl);
      setSigImgLocal((prev) => ({ ...prev, [loc]: dataUrl }));
    };
    reader.readAsDataURL(file);
  }

  function onClearLocalSig(loc: AppLocale) {
    clearSignatureImageDataUrl(loc);
    setSigImgLocal((prev) => ({ ...prev, [loc]: "" }));
  }

  function addContactAlias() {
    if (!model) return;
    const newAlias = { id: `c${Date.now()}`, name: "", email: "" };
    setModel({ ...model, contactAliases: [...(model.contactAliases || []), newAlias] });
  }

  function removeContactAlias(id: string) {
    if (!model) return;
    setModel({ ...model, contactAliases: (model.contactAliases || []).filter(c => c.id !== id) });
  }

  function updateContactAlias(id: string, field: "name" | "email", value: string) {
    if (!model) return;
    setModel({
      ...model,
      contactAliases: (model.contactAliases || []).map(c => c.id === id ? { ...c, [field]: value } : c)
    });
  }

  if (loading) {
    return <PanelState tone="loading" title="A carregar definições" description="Estamos a preparar as preferências guardadas deste utilizador." />;
  }

  if (!model) {
    return <PanelState tone="error" title="Não foi possível carregar as definições" description="Volta a abrir o painel ou tenta novamente dentro de instantes." />;
  }

  if (loading) {
    return <div style={S.note}>A carregar definições…</div>;
  }
  if (!model) {
    return <div style={S.error}>Não foi possível carregar as definições.</div>;
  }

  return (
    <div>
      <div style={S.headerRow}>
        <div style={S.hTitle}>{title}</div>
        <div style={{ display: "flex", gap: 8 }}>
          <button style={S.btnGhost} onClick={onReset} disabled={saving} title="Repor">
            <Icons.RotateCcw size={12} style={{ marginRight: "4px" }} />
            Repor
          </button>
          <button style={S.btn} onClick={onSave} disabled={saving}>
            <Icons.Save size={12} style={{ marginRight: "4px" }} />
            {saving ? "A guardar…" : "Guardar"}
          </button>
        </div>
      </div>

      <div style={S.card}>
        <div style={S.sidebar}>
          <button style={section === "general" ? S.sideItemOn : S.sideItem} onClick={() => setSection("general")}>
            Geral
          </button>
          <button style={section === "conns" ? S.sideItemOn : S.sideItem} onClick={() => setSection("conns")}>
            Ligações
          </button>
          <button style={section === "references" ? S.sideItemOn : S.sideItem} onClick={() => setSection("references")}>
            Referencias
          </button>
          <button style={section === "groups" ? S.sideItemOn : S.sideItem} onClick={() => setSection("groups")}>
            Grupos
          </button>
          <button style={section === "crm2layout" ? S.sideItemOn : S.sideItem} onClick={() => setSection("crm2layout")}>
            CRM2
          </button>
          <button style={section === "protection" ? S.sideItemOn : S.sideItem} onClick={() => setSection("protection")}>
            Proteção
          </button>
        </div>

        <div style={S.content}>
          {section === "conns" && (
            <ConnectionSettings
              model={model}
              setModel={setModel}
              setStatus={setStatus}
              availableModels={availableModels}
              fetchingModels={fetchingModels}
              refreshModels={fetchModels}
            />
          )}

          {section === "general" && (
            <div style={{ display: "grid", gap: 10 }}>
              <Field label="Idioma da app">
                <select
                  style={S.select}
                  value={model.appLanguage}
                  onChange={(e) => setModel({ ...model, appLanguage: e.target.value as AppLocale })}
                >
                  {Object.keys(LOCALE_LABEL).map((k) => (
                    <option key={k} value={k}>
                      {LOCALE_LABEL[k as AppLocale]}
                    </option>
                  ))}
                </select>
              </Field>

              <Field label="Tema (skin)">
                <select
                  style={S.select}
                  value={model.skinId || "classic"}
                  onChange={(e) => {
                    const v = e.target.value as SkinId;
                    setModel({ ...model, skinId: v });
                    try {
                      applySkin(v);
                    } catch {
                      /* ignore */
                    }
                  }}
                >
                  {SKIN_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
                <div style={S.hint}>
                  Classic mantém o visual atual. MailMaestro é compacto. Vibrant é o novo design Cockpit 3.0 com Glassmorphism.
                </div>
              </Field>

              <Field label="Idioma de leitura (resumo/rapidas)">
                <select
                  style={S.select}
                  value={model.readingLanguage}
                  onChange={(e) => setModel({ ...model, readingLanguage: e.target.value as LangOption })}
                >
                  {LANG_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </Field>

              <Field label="Idioma de resposta">
                <select
                  style={S.select}
                  value={model.replyLanguage}
                  onChange={(e) => setModel({ ...model, replyLanguage: e.target.value as LangOption })}
                >
                  {LANG_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </Field>

              <Field label="Idiomas no seletor rápido (barra inferior)">
                <div style={{ display: "flex", flexWrap: "wrap", gap: 10 }}>
                  {PICKER_LANGS.map((loc) => {
                    const enabled = (model.enabledLanguages && model.enabledLanguages.length > 0 ? model.enabledLanguages : PICKER_LANGS).includes(loc);
                    return (
                      <label
                        key={loc}
                        style={{
                          display: "inline-flex",
                          alignItems: "center",
                          gap: 6,
                          padding: "6px 10px",
                          borderRadius: 10,
                          border: "1px solid #e6e6e6",
                          background: enabled ? "#f7fbff" : "#fff",
                        }}
                      >
                        <input
                          type="checkbox"
                          checked={enabled}
                          onChange={(e) => {
                            const base = model.enabledLanguages && model.enabledLanguages.length > 0 ? [...model.enabledLanguages] : [...PICKER_LANGS];
                            const next = e.target.checked ? Array.from(new Set([...base, loc])) : base.filter((x) => x !== loc);
                            // keep at least one language visible
                            setModel({ ...model, enabledLanguages: next.length ? next : base });
                          }}
                        />
                        <span style={{ fontWeight: 700, fontSize: 12 }}>{localeShort(loc)}</span>
                        <span style={{ fontSize: 12, opacity: 0.8 }}>{LOCALE_LABEL[loc]}</span>
                      </label>
                    );
                  })}
                </div>
                <div style={{ ...S.hint, marginTop: 6 }}>Estas opções controlam o menu rápido de idiomas (ícone ao lado de “Resumo”).</div>
              </Field>

              <Field label="Tom">
                <select style={S.select} value={model.tone} onChange={(e) => setModel({ ...model, tone: e.target.value as any })}>
                  {TONE_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </Field>

              <Field label="Tamanho da resposta">
                <select
                  style={S.select}
                  value={model.length}
                  onChange={(e) => setModel({ ...model, length: e.target.value as ReplyLength })}
                >
                  {LENGTH_OPTIONS.map((o) => (
                    <option key={o.value} value={o.value}>
                      {o.label}
                    </option>
                  ))}
                </select>
              </Field>

              <div style={S.hint}>
                Nota: nesta fase, estas definições são a base. A IA vai começar a usá-las progressivamente (idioma/tom/tamanho).
              </div>
            </div>
          )}

          {section === "ai" && (
            <div style={{ display: "grid", gap: 16 }}>
              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.aiManualOnly !== false}
                  onChange={(e) => setModel({ ...model, aiManualOnly: e.target.checked })}
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>IA apenas manual</div>
                  <div style={S.hint}>Quando ativo, a app so usa IA por clique explicito. Quando desligado, voltam as analises automaticas onde existirem.</div>
                </div>
              </label>

              <div>
                <div style={S.fieldLabel}>Base de Conhecimento</div>
                <div style={{ ...S.hint, marginBottom: 8 }}>Identifica factos, regras ou dados da empresa que a IA deve saber (ex: NIF, IBAN, Morada).</div>
                <textarea
                  style={{ ...S.textarea, minHeight: 120 }}
                  value={(model.aiKnowledge || []).join("\n")}
                  onChange={(e) =>
                    setModel({ ...model, aiKnowledge: e.target.value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean) })
                  }
                  placeholder="Ex: NIF: 512345678&#10;IBAN: PT50 0000...&#10;Prazo de entrega: 48h"
                />
              </div>

              <div style={{ borderTop: "1px solid var(--iccc-card-border)", paddingTop: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={S.fieldLabel}>MODS / Response Presets</div>
                </div>
                <div style={{ ...S.hint, padding: 12, border: "1px dashed var(--iccc-card-border)", borderRadius: 12 }}>
                  Os MODS são geridos nas AI Settings. Esta secção usa a mesma fonte oficial, mas não edita `responsePresets`.
                </div>
              </div>

              <div style={{ borderTop: "1px solid var(--iccc-card-border)", paddingTop: 16 }}>
                <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: 8 }}>
                  <div style={S.fieldLabel}>Atalhos de Contactos (Reenvios)</div>
                  <button style={{ ...S.btnGhost, padding: "4px 10px", height: "auto" }} onClick={addContactAlias}>
                    <Icons.Settings size={12} style={{ marginRight: 4 }} />
                    Adicionar
                  </button>
                </div>
                <div style={{ ...S.hint, marginBottom: 12 }}>Mapeia nomes de fábricas ou entidades (ex: Ragno) para os seus emails.</div>

                <div style={{ display: "grid", gap: 10 }}>
                  {(model.contactAliases || []).map((c) => (
                    <div key={c.id} style={{ display: "flex", gap: 8, alignItems: "center" }}>
                      <input
                        style={{ ...S.input, flex: 2 }}
                        value={c.name}
                        onChange={(e) => updateContactAlias(c.id, "name", e.target.value)}
                        placeholder="Nome (ex: Ragno)"
                        title="Nome da entidade"
                      />
                      <input
                        style={{ ...S.input, flex: 3 }}
                        value={c.email}
                        onChange={(e) => updateContactAlias(c.id, "email", e.target.value)}
                        placeholder="Email (ex: info@ragno.it)"
                        title="Email da entidade"
                      />
                      <button
                        style={{ ...S.btnGhost, borderColor: "#fca5a5", color: "#ef4444", width: "32px", flexShrink: 0 }}
                        onClick={() => removeContactAlias(c.id)}
                        title="Remover atalho"
                      >
                        <Icons.Trash size={12} />
                      </button>
                    </div>
                  ))}
                  {(!model.contactAliases || model.contactAliases.length === 0) && (
                    <div style={{ ...S.hint, textAlign: "center", padding: 10, border: "1px dashed var(--iccc-card-border)", borderRadius: 12 }}>
                      Nenhum atalho criado.
                    </div>
                  )}
                </div>
              </div>
            </div>
          )}

          {section === "persona" && (
            <div style={{ display: "grid", gap: 12 }}>
              <div style={S.hint}>
                Define quem és e como escreves para que a IA possa imitar o teu estilo ("Ghost Writer").
              </div>

              <Field label="A minha função / Empresa">
                <input
                  style={S.input}
                  placeholder="Ex: Gestor de clientes na empresa"
                  value={model.userRole || ""}
                  onChange={(e) => setModel({ ...model, userRole: e.target.value })}
                />
              </Field>

              <Field label="Estilo e Contexto">
                <textarea
                  style={{ ...S.textarea, minHeight: 60 }}
                  placeholder="Ex: Escrevo de forma direta, saúdo sempre com 'Olá', não uso formalismos excessivos."
                  value={model.styleContext || ""}
                  onChange={(e) => setModel({ ...model, styleContext: e.target.value })}
                />
              </Field>

              <Field label="Exemplos de Escrita (Style Mimic)">
                <textarea
                  style={{ ...S.textarea, minHeight: 120 }}
                  placeholder="Cola aqui 2 ou 3 emails que escreveste no passado para a IA aprender o teu ritmo."
                  value={model.styleExamples || ""}
                  onChange={(e) => setModel({ ...model, styleExamples: e.target.value })}
                />
              </Field>

              <div style={{ ...S.fieldLabel, marginTop: 10 }}>Links de Reunião (Calendário)</div>
              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Microsoft Teams">
                  <input
                    style={S.input}
                    placeholder="Link da tua sala pessoal"
                    value={model.meetingLinks?.teams || ""}
                    onChange={(e) => setModel({ ...model, meetingLinks: { ...(model.meetingLinks || {}), teams: e.target.value } })}
                  />
                </Field>
                <Field label="Zoom">
                  <input
                    style={S.input}
                    placeholder="Link da tua sala pessoal"
                    value={model.meetingLinks?.zoom || ""}
                    onChange={(e) => setModel({ ...model, meetingLinks: { ...(model.meetingLinks || {}), zoom: e.target.value } })}
                  />
                </Field>
                <div style={{ gridColumn: "span 2" }}>
                  <Field label="Google Meet">
                    <input
                      style={S.input}
                      placeholder="Link da tua sala pessoal"
                      value={model.meetingLinks?.meet || ""}
                      onChange={(e) => setModel({ ...model, meetingLinks: { ...(model.meetingLinks || {}), meet: e.target.value } })}
                    />
                  </Field>
                </div>
              </div>
              <div style={S.hint}>Estes links serão usados automaticamente quando criares um agendamento via Cockpit.</div>
            </div>
          )}

          {section === "signature" && (
            <div style={{ display: "grid", gap: 12 }}>
              <div style={S.hint}>
                Assinatura por idioma. Podes usar <strong>Imagem</strong> (upload/URL), <strong>HTML</strong> (formatação) e/ou{" "}
                texto simples (fallback). A imagem enviada por upload é guardada <strong>localmente</strong>.
              </div>

              {(PICKER_LANGS as AppLocale[]).map((loc) => {
                const localImg = (sigImgLocal?.[loc] || "").trim();
                const urlImg = String(model.signatureImageUrl?.[loc] || "").trim();
                const maxW = Number(model.signatureImageMaxWidth?.[loc] ?? 260) || 260;
                const previewSrc = localImg || urlImg;

                return (
                  <div key={loc}>
                    <div style={S.fieldLabel}>{LOCALE_LABEL[loc]}</div>

                    {/* Image signature */}
                    <div style={{ display: "grid", gap: 8, marginBottom: 10, padding: 10, border: "1px solid #e6e8ef", borderRadius: 12 }}>
                      <div style={{ fontSize: 12, color: "#566" }}>
                        Assinatura <strong>Imagem</strong>
                      </div>

                      <div style={{ display: "flex", gap: 10, flexWrap: "wrap", alignItems: "center" }}>
                        <label style={{ display: "inline-flex", gap: 8, alignItems: "center" }}>
                          <span style={{ fontSize: 12, opacity: 0.85 }}>Upload:</span>
                          <input
                            type="file"
                            accept="image/*"
                            onChange={(e) => {
                              const f = e.target.files?.[0];
                              if (f) onUploadSig(loc, f);
                              // reset input so same file can be re-uploaded
                              (e.target as any).value = "";
                            }}
                          />
                        </label>

                        <button
                          style={S.btnGhost}
                          type="button"
                          onClick={() => onClearLocalSig(loc)}
                          disabled={!localImg}
                          title="Remove apenas a imagem guardada localmente (upload)"
                        >
                          Remover upload
                        </button>
                      </div>

                      <div style={{ display: "grid", gap: 6 }}>
                        <div style={{ fontSize: 12, color: "#566" }}>URL alternativa (se não quiseres upload)</div>
                        <input
                          style={S.input}
                          value={urlImg}
                          onChange={(e) => setSigUrl(loc, e.target.value)}
                          placeholder="https://.../assinatura.png"
                        />
                      </div>

                      <div style={{ display: "grid", gap: 6, maxWidth: 220 }}>
                        <div style={{ fontSize: 12, color: "#566" }}>Largura máx. (px)</div>
                        <input
                          style={S.input}
                          type="number"
                          min={120}
                          max={900}
                          value={String(maxW)}
                          onChange={(e) => setSigMaxW(loc, parseInt(e.target.value || "260", 10))}
                        />
                      </div>

                      {previewSrc ? (
                        <div style={{ marginTop: 6 }}>
                          <div style={{ fontSize: 11, color: "#66719a", marginBottom: 6 }}>Pré-visualização</div>
                          <div style={{ border: "1px dashed #d7dbeb", borderRadius: 12, padding: 10, background: "#fafbff" }}>
                            <img src={previewSrc} alt="" style={{ maxWidth: Math.max(120, Math.min(900, maxW)), height: "auto", display: "block" }} />
                          </div>
                          <div style={{ ...S.hint, marginTop: 6 }}>
                            Dica: mantém o ficheiro pequeno. Upload em dataURL pode ficar pesado (melhor PNG otimizado ou usar URL).
                          </div>
                        </div>
                      ) : (
                        <div style={S.hint}>Sem imagem configurada neste idioma.</div>
                      )}
                    </div>

                    {/* HTML signature */}
                    <div style={{ display: "grid", gap: 8, marginBottom: 8 }}>
                      <div style={{ fontSize: 12, color: "#566" }}>
                        Assinatura <strong>HTML</strong>
                      </div>
                      <textarea
                        style={S.textarea}
                        value={(model.signaturesHtml && model.signaturesHtml[loc]) || ""}
                        onChange={(e) =>
                          setModel({
                            ...model,
                            signaturesHtml: { ...(model.signaturesHtml || {}), [loc]: e.target.value },
                          })
                        }
                        placeholder='Ex.: <div>Com os melhores cumprimentos,<br/>Nome Apelido<br/>Empresa</div>'
                      />
                    </div>

                    {/* Text signature */}
                    <div style={{ display: "grid", gap: 8 }}>
                      <div style={{ fontSize: 12, color: "#566" }}>Assinatura (texto simples)</div>
                      <textarea
                        style={S.textarea}
                        value={model.signatures?.[loc] || ""}
                        onChange={(e) =>
                          setModel({
                            ...model,
                            signatures: { ...model.signatures, [loc]: e.target.value },
                          })
                        }
                        placeholder={"Ex.:\nCom os melhores cumprimentos,\nNome Apelido\nEmpresa"}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
          )}

          {section === "references" && (
            <div style={{ display: "grid", gap: 14 }}>
              <div style={S.hint}>
                Gera codigos de referencia para novos leads, projetos, tarefas e tickets criados pelo add-in. Os valores
                guardados aqui passam a ser a regra para futuros registos.
              </div>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.referenceCodes.enabled}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      referenceCodes: {
                        ...model.referenceCodes,
                        enabled: e.target.checked,
                      },
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Ativar codigos de referencia</div>
                  <div style={S.hint}>Os registos existentes nao sao alterados.</div>
                </div>
              </label>

              <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
                <Field label="Modo de numeracao">
                  <select
                    style={S.select}
                    value={model.referenceCodes.counterMode}
                    onChange={(e) =>
                      setModel({
                        ...model,
                        referenceCodes: {
                          ...model.referenceCodes,
                          counterMode: e.target.value as CockpitSettingsV1["referenceCodes"]["counterMode"],
                        },
                      })
                    }
                  >
                    <option value="per_type">Contador por tipo</option>
                    <option value="global">Contador global</option>
                  </select>
                </Field>

                <Field label="Posicao no titulo">
                  <select
                    style={S.select}
                    value={model.referenceCodes.position}
                    onChange={(e) =>
                      setModel({
                        ...model,
                        referenceCodes: {
                          ...model.referenceCodes,
                          position: e.target.value as CockpitSettingsV1["referenceCodes"]["position"],
                        },
                      })
                    }
                  >
                    <option value="prefix">No inicio</option>
                    <option value="suffix">No fim</option>
                  </select>
                </Field>
              </div>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.referenceCodes.includeYear}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      referenceCodes: {
                        ...model.referenceCodes,
                        includeYear: e.target.checked,
                      },
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Incluir ano no codigo</div>
                  <div style={S.hint}>O exemplo e o proximo codigo passam a incluir o ano atual.</div>
                </div>
              </label>

              <div style={{ display: "grid", gap: 10 }}>
                <div style={S.fieldLabel}>Prefixos por entidade</div>
                {(Object.keys(REFERENCE_ENTITY_LABELS) as ReferenceEntityKey[]).map((entity) => (
                  <div key={entity} style={S.referenceCard}>
                    <div style={{ display: "grid", gridTemplateColumns: "120px 1fr", gap: 10, alignItems: "center" }}>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>{REFERENCE_ENTITY_LABELS[entity]}</div>
                      <input
                        style={S.input}
                        value={model.referenceCodes.prefixes[entity] || ""}
                        onChange={(e) =>
                          setModel({
                            ...model,
                            referenceCodes: {
                              ...model.referenceCodes,
                              prefixes: {
                                ...model.referenceCodes.prefixes,
                                [entity]: e.target.value,
                              },
                            },
                          })
                        }
                        placeholder="Prefixo opcional"
                      />
                    </div>
                    <div style={{ ...S.hint, marginTop: 8 }}>Pre-visualizacao: {previewReferenceCode(model, entity)}</div>
                    <div style={{ ...S.hint, marginTop: 4 }}>
                      Contador atual: {model.referenceCodes.counterMode === "global"
                        ? model.referenceCodes.counters.global
                        : model.referenceCodes.counters.perType[entity]}
                    </div>
                  </div>
                ))}
              </div>

              <div style={{ display: "grid", gap: 10 }}>
                <div style={S.fieldLabel}>Reset de contadores</div>
                <div style={{ display: "flex", flexWrap: "wrap", gap: 8 }}>
                  <button
                    type="button"
                    style={S.btnGhost}
                    onClick={() =>
                      setModel({
                        ...model,
                        referenceCodes: {
                          ...model.referenceCodes,
                          counters: {
                            ...model.referenceCodes.counters,
                            global: 0,
                          },
                        },
                      })
                    }
                  >
                    Reset global
                  </button>
                  {(Object.keys(REFERENCE_ENTITY_LABELS) as ReferenceEntityKey[]).map((entity) => (
                    <button
                      key={entity}
                      type="button"
                      style={S.btnGhost}
                      onClick={() =>
                        setModel({
                          ...model,
                          referenceCodes: {
                            ...model.referenceCodes,
                            counters: {
                              ...model.referenceCodes.counters,
                              perType: {
                                ...model.referenceCodes.counters.perType,
                                [entity]: 0,
                              },
                            },
                          },
                        })
                      }
                    >
                      Reset {REFERENCE_ENTITY_LABELS[entity]}
                    </button>
                  ))}
                </div>
                <div style={S.hint}>O reset afeta apenas futuros codigos.</div>
              </div>
            </div>
          )}

          {section === "groups" && (
            <div style={{ display: "grid", gap: 14 }}>
              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.groupLabelsManagerEnabled !== false}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupLabelsManagerEnabled: e.target.checked,
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Ativar gestor de etiquetas</div>
                  <div style={S.hint}>Liga a gestão central de etiquetas dentro da aba Grupos. Quando desligado, a app esconde o gestor dedicado de etiquetas.</div>
                </div>
              </label>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.groupTicketsEnabled !== false}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupTicketsEnabled: e.target.checked,
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Ativar tickets nos grupos</div>
                  <div style={S.hint}>Liga o extra de tickets dentro da aba Grupos. As séries, contadores e regras de ligação ficam na roda dentada da própria aba.</div>
                </div>
              </label>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.groupOutlookCategories?.enabled === true}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupOutlookCategories: {
                        enabled: e.target.checked,
                        includeGroups: model.groupOutlookCategories?.includeGroups !== false,
                        includeTickets: model.groupOutlookCategories?.includeTickets !== false,
                        includeStatuses: model.groupOutlookCategories?.includeStatuses !== false,
                        includeLabels: model.groupOutlookCategories?.includeLabels === true,
                      },
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Escrever categorias no Outlook</div>
                  <div style={S.hint}>Aplica categorias automáticas no email atual com base em grupos, tickets e estado.</div>
                </div>
              </label>

              {model.groupOutlookCategories?.enabled === true && (
                <div style={{ ...S.referenceCard, display: "grid", gap: 10 }}>
                  <div style={S.fieldLabel}>Categorias automáticas</div>
                  <label style={S.toggleRow}>
                    <input
                      type="checkbox"
                      checked={model.groupOutlookCategories?.includeGroups !== false}
                      onChange={(e) =>
                        setModel({
                          ...model,
                          groupOutlookCategories: {
                            ...model.groupOutlookCategories,
                            enabled: true,
                            includeGroups: e.target.checked,
                          },
                        })
                      }
                    />
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Grupo</div>
                      <div style={S.hint}>Exemplo: <b>Grupo: Encomendas</b></div>
                    </div>
                  </label>

                  <label style={S.toggleRow}>
                    <input
                      type="checkbox"
                      checked={model.groupOutlookCategories?.includeTickets !== false}
                      onChange={(e) =>
                        setModel({
                          ...model,
                          groupOutlookCategories: {
                            ...model.groupOutlookCategories,
                            enabled: true,
                            includeTickets: e.target.checked,
                          },
                        })
                      }
                    />
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Ticket</div>
                      <div style={S.hint}>Exemplo: <b>TK: RTK-26-0001</b></div>
                    </div>
                  </label>

                  <label style={S.toggleRow}>
                    <input
                      type="checkbox"
                      checked={model.groupOutlookCategories?.includeStatuses !== false}
                      onChange={(e) =>
                        setModel({
                          ...model,
                          groupOutlookCategories: {
                            ...model.groupOutlookCategories,
                            enabled: true,
                            includeStatuses: e.target.checked,
                          },
                        })
                      }
                    />
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Estado</div>
                      <div style={S.hint}>Exemplo: <b>Gr: Em analise</b> · <b>E-Tk: Aberto</b> · <b>E-Et: Em progresso</b></div>
                    </div>
                  </label>

                  <label style={S.toggleRow}>
                    <input
                      type="checkbox"
                      checked={model.groupOutlookCategories?.includeLabels === true}
                      onChange={(e) =>
                        setModel({
                          ...model,
                          groupOutlookCategories: {
                            ...model.groupOutlookCategories,
                            enabled: true,
                            includeLabels: e.target.checked,
                          },
                        })
                      }
                    />
                    <div>
                      <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Etiquetas</div>
                      <div style={S.hint}>Exemplo: <b>PED-1128</b></div>
                    </div>
                  </label>
                </div>
              )}

              <PanelState
                compact
                tone={
                  validatingGroupStorage
                    ? "loading"
                    : groupStorageValidation?.supported === false
                      ? "warning"
                      : "info"
                }
                title="Politica executavel de storage desta fase"
                description={
                  validatingGroupStorage
                    ? "A validar o destino de gravacao no host atual do servidor."
                    : groupStorageValidation?.supported === false
                      ? groupStorageValidation.blockingReason || "O destino configurado ainda nao fecha de forma executavel nesta arquitetura."
                      : "A persistencia final continua na app (`/api/links/*`). Os modos file-backed passam a depender de validacao real do caminho e nao apenas de labels."
                }
              />

              <div style={S.referenceCard}>
                <div style={S.fieldLabel}>Etiquetas conhecidas</div>
                <div style={{ display: "grid", gap: 6 }}>
                  <div style={S.hint}>
                    Estado: <b>{model.groupLabelsManagerEnabled !== false ? "Ativo" : "Desativado"}</b>
                  </div>
                  <div style={S.hint}>
                    Catálogo atual: <b>{(model.groupLabelCatalog || []).length}</b> etiqueta(s)
                  </div>
                  <div style={S.hint}>
                    A gestão diária, renomeação e limpeza das etiquetas passa a ser feita diretamente na aba Grupos pela roda dentada.
                  </div>
                  <div style={S.hint}>
                    Tickets: <b>{model.groupTicketsEnabled !== false ? "Ativos" : "Desativados"}</b>
                  </div>
                  <div style={S.hint}>
                    Categorias Outlook: <b>{model.groupOutlookCategories?.enabled === true ? "Ativas" : "Desativadas"}</b>
                  </div>
                </div>
              </div>

              <Field label="Modo de armazenamento">
                <select
                  style={S.select}
                  value={model.groupStorage.mode}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupStorage: {
                        ...model.groupStorage,
                        mode: e.target.value as GroupStorageMode,
                      },
                    })
                  }
                >
                  {EXECUTABLE_GROUP_STORAGE_MODE_OPTIONS.map((entry) => (
                    <option key={entry.value} value={entry.value}>
                      {entry.label}
                    </option>
                  ))}
                </select>
                <div style={S.hint}>
                  Todos os modos continuam visiveis no contrato, mas os file-backed so ficam validos quando o caminho e realmente acessivel ao servidor.
                </div>
                <div style={S.hint}>
                  `local_device` significa path local/UNC visivel para o host do servidor; nao representa automaticamente o disco do utilizador sem bridge nativa.
                </div>
              </Field>

              {model.groupStorage.mode === "hybrid" ? (
                <Field label="Primario do modo hibrido">
                  <select
                    style={S.select}
                    value={model.groupStorage.hybrid.primaryTarget}
                    onChange={(e) =>
                      setModel({
                        ...model,
                        groupStorage: {
                          ...model.groupStorage,
                          hybrid: {
                            ...model.groupStorage.hybrid,
                            primaryTarget: e.target.value === "local_device" ? "local_device" : "chosen_folder",
                          },
                        },
                      })
                    }
                  >
                    <option value="chosen_folder">Pasta escolhida</option>
                    <option value="local_device">Local acessivel ao servidor</option>
                  </select>
                </Field>
              ) : null}

              {model.groupStorage.mode !== "supabase" && model.groupStorage.mode !== "local_device" ? (
                <Field label="Tipo de destino escolhido">
                  <input
                    style={{ ...S.input, background: "rgba(248,250,252,0.94)" }}
                    value="Pasta fisica / sincronizada"
                    disabled
                  />
                  <div style={S.hint}>
                    Nesta fase sem Graph/admin, `chosen_folder` e `hybrid` aceitam apenas pasta fisica, pasta sincronizada local ou caminho UNC validado no servidor.
                  </div>
                </Field>
              ) : null}

              <Field label={model.groupStorage.mode === "local_device" ? "Destino local acessivel ao servidor" : "Destino principal"}>
                <input
                  style={S.input}
                  value={
                    model.groupStorage.mode === "local_device"
                      ? model.groupStorage.localDevice.rootPath || ""
                      : model.groupStorage.mode === "supabase"
                        ? ""
                        : model.groupStorage.chosenFolder.path || model.groupStorage.baseFolderPath || ""
                  }
                  disabled={!groupStoragePathRequired}
                  onChange={(e) =>
                    setModel({
                      ...model,
                        groupStorage: {
                          ...model.groupStorage,
                          baseFolderPath: e.target.value,
                          chosenFolder: {
                            ...model.groupStorage.chosenFolder,
                            path: e.target.value,
                            kind: "filesystem",
                          },
                          localDevice: {
                            ...model.groupStorage.localDevice,
                            rootPath:
                              model.groupStorage.mode === "local_device" || model.groupStorage.hybrid.primaryTarget === "local_device"
                              ? e.target.value
                              : model.groupStorage.localDevice.rootPath,
                        },
                      },
                    })
                  }
                  placeholder={groupStoragePathRequired ? "C:\\Documentos\\InboxCockpit\\Grupos ou \\\\servidor\\partilha\\Grupos" : "Nao usado no modo Cockpit Cloud"}
                />
                <div style={S.hint}>
                  Usa apenas <b>caminho local do host, pasta sincronizada local ou UNC</b>. URL web fica auditada e bloqueada explicitamente se o backend nao a suportar.
                </div>
              </Field>

              <div style={{ display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                <button
                  style={S.btnGhost}
                  onClick={() => model && void validateCurrentGroupStorage(model)}
                  disabled={validatingGroupStorage}
                >
                  {validatingGroupStorage ? "A validar..." : "Validar destino"}
                </button>
                <div style={S.hint}>
                  Picker real de pasta continua bloqueado pelo host atual; a alternativa suportada nesta arquitetura e path manual + validacao real no servidor.
                </div>
              </div>

              {groupStorageValidation ? (
                <div style={S.referenceCard}>
                  <div style={S.fieldLabel}>Validacao real do destino</div>
                  <div style={{ display: "grid", gap: 6 }}>
                    <div style={S.hint}>
                      Estado: <b>{groupStorageValidation.supported ? "Suportado neste host" : "Bloqueado neste host"}</b>
                    </div>
                    <div style={S.hint}>
                      Provider efetivo: <b>{groupStorageValidation.provider}</b>
                    </div>
                    <div style={S.hint}>
                      Caminho normalizado: <b>{groupStorageValidation.normalizedBasePath || "n/a"}</b>
                    </div>
                    <div style={S.hint}>
                      Manifesto espelhado: <b>{groupStorageValidation.canStoreManifest ? "Sim" : "Nao"}</b>
                    </div>
                    <div style={S.hint}>
                      Binario real: <b>{groupStorageValidation.canStoreBinary ? "Sim" : "Nao"}</b>
                    </div>
                    {groupStorageValidation.blockingReason ? (
                      <div style={S.hint}>
                        Bloqueio: <b>{groupStorageValidation.blockingReason}</b>
                      </div>
                    ) : null}
                    {groupStorageValidation.requiredChange ? (
                      <div style={S.hint}>
                        Mudanca minima necessaria: <b>{groupStorageValidation.requiredChange}</b>
                      </div>
                    ) : null}
                    {groupStorageValidation.pickerBlockedReason ? (
                      <div style={S.hint}>
                        Picker/path: <b>{groupStorageValidation.pickerBlockedReason}</b>
                      </div>
                    ) : null}
                  </div>
                </div>
              ) : null}

              {groupStorageValidation?.architecturalBlocker === "web_document_library_requires_graph_backend"
                || groupStorageRequiresGraphOrAdmin ? (
                <div style={S.referenceCard}>
                  <div style={S.fieldLabel}>Prova tecnica do bloqueio da URL web</div>
                  <div style={{ display: "grid", gap: 6 }}>
                    {GROUP_STORAGE_WEB_URL_BLOCKER_FACTS.map((fact) => (
                      <div key={fact} style={S.hint}>
                        - {fact}
                      </div>
                    ))}
                    <div style={S.hint}>
                      Resultado operacional desta fase: URL web fica explicitamente fora de scope e nao entra no runtime executavel de `local_device`, `chosen_folder` ou `hybrid`.
                    </div>
                  </div>
                </div>
              ) : null}

              <Field label="Limiar para decisao de anexos (MB)">
                <input
                  style={S.input}
                  type="number"
                  min={1}
                  max={250}
                  value={model.groupStorage.attachmentPromptThresholdMb || 10}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupStorage: {
                        ...model.groupStorage,
                        attachmentPromptThresholdMb: Number(e.target.value || 10) || 10,
                      },
                    })
                  }
                />
                <div style={S.hint}>
                  Metadata do anexo sobe sempre com o email classificado; binario fora do Cockpit Cloud fica best-effort e continua a pedir caminho real suportado.
                </div>
              </Field>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.groupStorage.autoCreateFolderOnGroupCreate}
                  disabled
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupStorage: {
                        ...model.groupStorage,
                        autoCreateFolderOnGroupCreate: e.target.checked,
                      },
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Criar pasta automaticamente ao criar grupo</div>
                  <div style={S.hint}>Mantido apenas como legado de configuracao. Esta automatizacao ainda nao fica executavel nesta fase.</div>
                </div>
              </label>

              <label style={S.toggleRow}>
                <input
                  type="checkbox"
                  checked={model.groupStorage.ignoreInlineAttachments}
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupStorage: {
                        ...model.groupStorage,
                        ignoreInlineAttachments: e.target.checked,
                      },
                    })
                  }
                />
                <div>
                  <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Ignorar anexos inline e imagens de assinatura</div>
                  <div style={S.hint}>Esta guarda continua real e afeta a preparacao/persistencia de anexos para reduzir ruido.</div>
                </div>
              </label>

              <Field label="Viewer preferido">
                <select
                  style={S.select}
                  value={model.groupStorage.suggestedViewer}
                  disabled
                  onChange={(e) =>
                    setModel({
                      ...model,
                      groupStorage: {
                        ...model.groupStorage,
                        suggestedViewer: e.target.value as "system" | "inline",
                      },
                    })
                  }
                >
                  <option value="inline">Viewer interno do cockpit</option>
                  <option value="system">Aplicação do sistema</option>
                </select>
                <div style={S.hint}>
                  Preferencia mantida apenas como legado local. Ainda nao existe caminho executavel fechado que a use de ponta a ponta nesta fase.
                </div>
              </Field>

              <div style={S.referenceCard}>
                <div style={S.fieldLabel}>Resumo executavel desta fase</div>
                <div style={{ display: "grid", gap: 6 }}>
                  <div style={S.hint}>
                    Modo escolhido: <b>{resolvedGroupStorage?.modeLabel || GROUP_STORAGE_MODE_LABELS[model.groupStorage.mode]}</b>
                  </div>
                  <div style={S.hint}>
                    Persistencia final: <b>sempre na app via /api/links/*</b>
                  </div>
                  <div style={S.hint}>
                    Binario de anexos/documentos: <b>{model.groupStorage.mode === "supabase" ? "Cockpit Cloud quando o payload traz conteudo" : groupStorageValidation?.supported ? "path validado no host atual + fallback controlado para metadata/cloud" : "bloqueado ate existir path fisico validado no servidor"}</b>
                  </div>
                  <div style={S.hint}>
                    Persistencia intermedia: <b>{model.groupsTabSettings.storageMode === "disabled" ? "desativada" : "IndexedDB local do add-in, com namespace logico"}</b>
                  </div>
                  <div style={S.hint}>
                    Sessao/cache: <b>prepareSession, seeds temporarios e fallback em memoria</b>
                  </div>
                </div>
              </div>

              <div style={S.referenceCard}>
                <div style={S.fieldLabel}>Tabela real de modos</div>
                <div style={{ display: "grid", gap: 8 }}>
                  {groupStorageCapabilities.map((entry) => (
                    <div key={entry.mode} style={{ display: "grid", gap: 3, padding: "8px 10px", borderRadius: 10, border: "1px solid var(--iccc-card-border)", background: "#fff" }}>
                      <div style={S.hint}>
                        <b>{entry.label}</b> - {entry.supported ? "Suportado / validado" : "Bloqueado ou dependente de validacao"}
                      </div>
                      <div style={S.hint}>Como grava: <b>{entry.howItWrites}</b></div>
                      <div style={S.hint}>Onde grava: <b>{entry.whereItWrites}</b></div>
                      <div style={S.hint}>Limitacao real: <b>{entry.limitations}</b></div>
                    </div>
                  ))}
                </div>
              </div>
            </div>
          )}

          {section === "crm2layout" && (
            <Crm2LayoutSettings
              model={model}
              setModel={setModel}
            />
          )}

          {section === ("protection" as any) && (
            <ProtectionSettings />
          )}

          {normalizeStatus(status) && (
            <PanelState
              tone={normalizeStatus(status)!.tone}
              title={normalizeStatus(status)!.title}
              description={normalizeStatus(status)!.description}
              compact
            />
          )}
        </div>
      </div>
    </div>
  );
}

function Field({ label, children }: { label: string; children: React.ReactNode }) {
  return (
    <div>
      <div style={S.fieldLabel}>{label}</div>
      {children}
    </div>
  );
}

function normalizeStatus(status: StatusValue): StatusNotice | null {
  if (!status) return null;
  if (typeof status !== "string") return status;
  if (/falha|erro/i.test(status)) {
    return { tone: "error", title: "Falha nas definições", description: status };
  }
  return { tone: "success", title: status, description: undefined };
}

const S: Record<string, React.CSSProperties> = {
  headerRow: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: 12,
    marginBottom: 10,
  },
  hTitle: { fontWeight: 800, fontSize: 14, color: "var(--iccc-text)" },

  card: {
    borderRadius: "var(--iccc-radius-card)",
    background: "var(--iccc-card-bg)",
    border: "1px solid var(--iccc-card-border)",
    boxShadow: "var(--iccc-shadow)",
    padding: 10,
    display: "grid",
    gridTemplateColumns: "110px 1fr",
    gap: 10,
    backdropFilter: "var(--iccc-glass-blur)",
    WebkitBackdropFilter: "var(--iccc-glass-blur)",
  },
  sidebar: {
    display: "grid",
    gap: 6,
    alignContent: "start",
  },
  content: {
    minHeight: 220,
    color: "var(--iccc-text)",
  },

  sideItem: {
    borderRadius: 8,
    padding: "6px 10px",
    border: "1px solid transparent",
    background: "transparent",
    fontSize: "11px",
    textAlign: "left",
    cursor: "pointer",
    color: "var(--iccc-text-muted)",
  },
  sideItemOn: {
    borderRadius: 8,
    padding: "6px 10px",
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(0,0,0,0.03)",
    fontSize: "11px",
    textAlign: "left",
    cursor: "pointer",
    color: "var(--iccc-pill-active-bg)",
    fontWeight: 700,
  },

  fieldLabel: {
    fontSize: 11,
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.02em",
    color: "var(--iccc-text-muted)",
    marginBottom: 6,
  },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(255,255,255,0.05)",
    color: "var(--iccc-text)",
    padding: "8px 10px",
    fontSize: 12,
    outline: "none",
  },
  input: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(255,255,255,0.05)",
    color: "var(--iccc-text)",
    padding: "8px 10px",
    fontSize: 12,
    outline: "none",
  },
  textarea: {
    width: "100%",
    minHeight: 80,
    borderRadius: 12,
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(255,255,255,0.05)",
    color: "var(--iccc-text)",
    padding: 10,
    fontSize: 12,
    outline: "none",
    resize: "vertical",
  },
  hint: {
    fontSize: 11,
    color: "var(--iccc-text-muted)",
    lineHeight: 1.35,
  },
  toggleRow: {
    display: "flex",
    alignItems: "flex-start",
    gap: 10,
    padding: "10px 12px",
    borderRadius: 12,
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(255,255,255,0.03)",
  },
  referenceCard: {
    padding: 12,
    borderRadius: 12,
    border: "1px solid var(--iccc-card-border)",
    background: "rgba(255,255,255,0.03)",
  },

  btn: {
    borderRadius: 999,
    border: "none",
    background: "var(--iccc-btn-bg)",
    color: "var(--iccc-btn-text)",
    padding: "6px 14px",
    fontSize: 11,
    fontWeight: 800,
    textTransform: "uppercase",
    cursor: "pointer",
    boxShadow: "0 4px 12px rgba(0,0,0,0.1)",
  },
  btnGhost: {
    borderRadius: 999,
    border: "1px solid var(--iccc-card-border)",
    background: "transparent",
    color: "var(--iccc-text)",
    padding: "6px 14px",
    fontSize: 11,
    fontWeight: 800,
    textTransform: "uppercase",
    cursor: "pointer",
  },

  okBox: {
    marginTop: 10,
    borderRadius: 12,
    padding: 10,
    fontSize: 11,
    fontWeight: 600,
    border: "1px solid var(--iccc-pill-active-bg)",
    background: "rgba(16, 185, 129, 0.1)",
    color: "var(--iccc-pill-active-bg)",
  },
  errorBox: {
    marginTop: 10,
    borderRadius: 12,
    padding: 10,
    fontSize: 11,
    fontWeight: 600,
    border: "1px solid #ef4444",
    background: "rgba(239, 68, 68, 0.1)",
    color: "#ef4444",
  },
  note: { fontSize: 11, color: "var(--iccc-text-muted)" },
  error: { fontSize: 11, color: "#ef4444" },
};

function Crm2LayoutSettings({
  model,
  setModel,
}: {
  model: CockpitSettingsV1;
  setModel: React.Dispatch<React.SetStateAction<CockpitSettingsV1 | null>>;
}) {
  const layout = model.crm2OdooLayout;
  const pdfGuideHref = "/docs/inbox-cockpit-crm2-odoo-studio-setup.pdf";
  const [layoutTarget, setLayoutTarget] = useState<Crm2OdooLayoutTarget>("project");
  const [isValidating, setIsValidating] = useState(false);
  const [validationError, setValidationError] = useState<string | null>(null);
  const [validation, setValidation] = useState<Crm2LayoutValidationResult | null>(null);
  const targetConfig = layout[layoutTarget];
  const targetMeta = {
    project: {
      singular: "Projeto",
      plural: "projetos",
      button: "Projetos",
      fixedInfoField: "x_studio_iccc_project_brief",
      historyField: "x_studio_iccc_project_history",
      documentsField: "x_studio_iccc_project_documents",
    },
    lead: {
      singular: "Lead",
      plural: "leads",
      button: "Leads",
      fixedInfoField: "x_studio_iccc_lead_brief",
      historyField: "x_studio_iccc_lead_history",
      documentsField: "x_studio_iccc_lead_documents",
    },
    task: {
      singular: "Tarefa",
      plural: "tarefas",
      button: "Tarefas",
      fixedInfoField: "x_studio_iccc_task_brief",
      historyField: "x_studio_iccc_task_history",
      documentsField: "x_studio_iccc_task_documents",
    },
    ticket: {
      singular: "Ticket",
      plural: "tickets",
      button: "Tickets",
      fixedInfoField: "x_studio_iccc_ticket_brief",
      historyField: "x_studio_iccc_ticket_history",
      documentsField: "x_studio_iccc_ticket_documents",
    },
  } as const;
  const targetLabel = targetMeta[layoutTarget].singular;
  const targetPluralLabel = targetMeta[layoutTarget].plural;
  const targetGuideDefaults = targetMeta[layoutTarget];
  const targetModeLabel = targetConfig.mode === "structured_project" ? `${targetLabel} com campos/abas proprias` : "Descricao apenas";

  function updateLayout<K extends keyof CockpitSettingsV1["crm2OdooLayout"]>(
    key: K,
    value: CockpitSettingsV1["crm2OdooLayout"][K],
  ) {
    setModel((prev) =>
      prev
        ? {
            ...prev,
            crm2OdooLayout: {
              ...prev.crm2OdooLayout,
              [key]: value,
            },
          }
        : prev,
    );
  }

  function updateTargetConfig<K extends keyof CockpitSettingsV1["crm2OdooLayout"]["project"]>(
    key: K,
    value: CockpitSettingsV1["crm2OdooLayout"]["project"][K],
  ) {
    setModel((prev) =>
      prev
        ? {
            ...prev,
            crm2OdooLayout: {
              ...prev.crm2OdooLayout,
              [layoutTarget]: {
                ...prev.crm2OdooLayout[layoutTarget],
                [key]: value,
              },
            },
          }
        : prev,
    );
  }

  async function runValidation() {
    setIsValidating(true);
    setValidationError(null);
    try {
      const result = await validateCrm2OdooLayout(layout, layoutTarget);
      setValidation(result);
    } catch (error: any) {
      setValidation(null);
      setValidationError(error?.message || "Nao foi possivel validar a configuracao no Odoo.");
    } finally {
      setIsValidating(false);
    }
  }

  return (
    <div style={{ display: "grid", gap: 14 }}>
      <PanelState
        compact
        tone="info"
        title="Estrategia de escrita do CRM2 no Odoo"
        description="Define por entidade se o CRM2 escreve apenas na descricao base ou se usa um layout estruturado com campos e abas preparados no Odoo Studio."
      />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10, flexWrap: "wrap" }}>
        <div style={S.hint}>
          Valida campos, tipos e presenca na vista form do modelo alvo antes de ativares o modo estruturado em producao.
        </div>
        <button
          type="button"
          style={isValidating ? { ...S.btnGhost, opacity: 0.7, cursor: "wait" } : S.btnGhost}
          onClick={runValidation}
          disabled={isValidating}
        >
          {isValidating ? "A validar..." : "Validar configuracao Odoo"}
        </button>
      </div>

      <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
        {(["project", "lead", "task", "ticket"] as Crm2OdooLayoutTarget[]).map((target) => (
          <button
            key={target}
            type="button"
            style={layoutTarget === target ? S.btn : S.btnGhost}
            onClick={() => {
              setLayoutTarget(target);
              setValidation(null);
              setValidationError(null);
            }}
          >
            {targetMeta[target].button}
          </button>
        ))}
      </div>

      <Field label={`Modo de escrita para ${targetPluralLabel}`}>
        <select
          style={S.select}
          value={targetConfig.mode}
          onChange={(e) => updateTargetConfig("mode", e.target.value === "structured_project" ? "structured_project" : "description_only")}
        >
          <option value="description_only">Descricao apenas (fallback universal)</option>
          <option value="structured_project">Layout estruturado com campos/abas proprias</option>
        </select>
        <div style={S.hint}>
          Esta escolha e independente por entidade. Podes deixar {targetPluralLabel} em "Descricao apenas" e usar modo estruturado noutros tipos ao mesmo tempo.
        </div>
      </Field>

      <div style={S.referenceCard}>
        <div style={S.fieldLabel}>Resumo de independencia</div>
        <div style={{ display: "grid", gap: 6 }}>
          {(["project", "lead", "task", "ticket"] as Crm2OdooLayoutTarget[]).map((target) => (
            <div key={target} style={S.hint}>
              {targetMeta[target].singular}: <b>{layout[target].mode === "structured_project" ? "Estruturado" : "Descricao apenas"}</b>
            </div>
          ))}
        </div>
      </div>

      <label style={S.toggleRow}>
        <input
          type="checkbox"
          checked={layout.includeAnchorIndex}
          onChange={(e) => updateLayout("includeAnchorIndex", e.target.checked)}
        />
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Criar indice de emails/posts</div>
          <div style={S.hint}>No modo estruturado, o historico pode abrir com um resumo navegavel dos emails da conversa.</div>
        </div>
      </label>

      <label style={S.toggleRow}>
        <input
          type="checkbox"
          checked={layout.showBackToTopLinks}
          onChange={(e) => updateLayout("showBackToTopLinks", e.target.checked)}
        />
        <div>
          <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Mostrar links "Voltar ao topo"</div>
          <div style={S.hint}>Ajuda a navegar historicos longos quando estivermos a escrever resumos e blocos por email no Odoo.</div>
        </div>
      </label>

      <div style={S.referenceCard}>
        <div style={S.fieldLabel}>Perfil estruturado recomendado para {targetPluralLabel}</div>
        <div style={{ display: "grid", gap: 10 }}>
          <Field label="Modelo Odoo alvo">
            <input
              style={{ ...S.input, background: "rgba(0,0,0,0.03)" }}
              value={targetConfig.model}
              readOnly
            />
          </Field>

          <Field label="Campo base da descricao">
            <input
              style={S.input}
              value={targetConfig.descriptionField}
              onChange={(e) => updateTargetConfig("descriptionField", e.target.value.trim())}
              placeholder="description"
            />
          </Field>

          <Field label="Campo de informacao fixa">
            <input
              style={S.input}
              value={targetConfig.fixedInfoField}
              onChange={(e) => updateTargetConfig("fixedInfoField", e.target.value.trim())}
              placeholder={targetGuideDefaults.fixedInfoField}
            />
          </Field>

          <Field label="Campo de historico">
            <input
              style={S.input}
              value={targetConfig.historyField}
              onChange={(e) => updateTargetConfig("historyField", e.target.value.trim())}
              placeholder={targetGuideDefaults.historyField}
            />
          </Field>

          <Field label="Campo de documentos">
            <input
              style={S.input}
              value={targetConfig.documentsField}
              onChange={(e) => updateTargetConfig("documentsField", e.target.value.trim())}
              placeholder={targetGuideDefaults.documentsField}
            />
          </Field>

          <div style={{ display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))", gap: 10 }}>
            <Field label="Tab: informacao fixa">
              <input
                style={S.input}
                value={targetConfig.fixedInfoTabLabel}
                onChange={(e) => updateTargetConfig("fixedInfoTabLabel", e.target.value)}
                placeholder="Informacao fixa"
              />
            </Field>

            <Field label="Tab: historico">
              <input
                style={S.input}
                value={targetConfig.historyTabLabel}
                onChange={(e) => updateTargetConfig("historyTabLabel", e.target.value)}
                placeholder="Historico"
              />
            </Field>

            <Field label="Tab: documentos">
              <input
                style={S.input}
                value={targetConfig.documentsTabLabel}
                onChange={(e) => updateTargetConfig("documentsTabLabel", e.target.value)}
                placeholder="Documentos"
              />
            </Field>
          </div>

          <label style={S.toggleRow}>
            <input
              type="checkbox"
              checked={targetConfig.fallbackToDescription}
              onChange={(e) => updateTargetConfig("fallbackToDescription", e.target.checked)}
            />
            <div>
              <div style={{ fontSize: 12, fontWeight: 700, color: "var(--iccc-text)" }}>Fallback automatico para descricao</div>
              <div style={S.hint}>Se faltar algum campo customizado no cliente, o CRM2 continua a funcionar e escreve no campo base da descricao de {targetPluralLabel}.</div>
            </div>
          </label>
        </div>
      </div>

      <div style={S.referenceCard}>
        <div style={S.fieldLabel}>Resumo atual</div>
        <div style={{ display: "grid", gap: 6 }}>
          <div style={S.hint}>
            Modo ativo: <b>{targetModeLabel}</b>
          </div>
          <div style={S.hint}>
            Modelo alvo: <b>{targetConfig.model}</b>
          </div>
          <div style={S.hint}>
            Campo base: <b>{targetConfig.descriptionField || "Por definir"}</b>
          </div>
          <div style={S.hint}>
            Estrutura recomendada: <b>{targetConfig.fixedInfoField || "?"}</b> / <b>{targetConfig.historyField || "?"}</b> / <b>{targetConfig.documentsField || "?"}</b>
          </div>
        </div>
      </div>

      <div style={S.referenceCard}>
        <div style={S.fieldLabel}>Checklist Odoo Studio</div>
        <div style={{ display: "grid", gap: 8 }}>
          <div style={S.hint}><b>1.</b> Abrir o Studio num registo de {targetLabel.toLowerCase()} e editar a vista de <b>{targetConfig.model}</b>.</div>
          <div style={S.hint}><b>2.</b> Confirmar que o campo base <b>{targetConfig.descriptionField || "description"}</b> existe e fica visivel no formulario.</div>
          <div style={S.hint}><b>3.</b> Criar o campo <b>{targetConfig.fixedInfoField || targetGuideDefaults.fixedInfoField}</b> para informacao fixa, de preferencia HTML.</div>
          <div style={S.hint}><b>4.</b> Criar o campo <b>{targetConfig.historyField || targetGuideDefaults.historyField}</b> para historico e o campo <b>{targetConfig.documentsField || targetGuideDefaults.documentsField}</b> para documentos.</div>
          <div style={S.hint}><b>5.</b> Adicionar as abas <b>{targetConfig.fixedInfoTabLabel || "Informacao fixa"}</b>, <b>{targetConfig.historyTabLabel || "Historico"}</b> e <b>{targetConfig.documentsTabLabel || "Documentos"}</b>, colocando um campo por aba.</div>
          <div style={S.hint}><b>6.</b> Guardar o Studio, voltar ao cockpit e correr <b>Validar configuracao Odoo</b>.</div>
        </div>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap", marginTop: 12 }}>
          <a href={pdfGuideHref} target="_blank" rel="noreferrer" style={{ ...S.btnGhost, textDecoration: "none", display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            Abrir guia PDF
          </a>
          <a href={pdfGuideHref} download style={{ ...S.btnGhost, textDecoration: "none", display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
            Descarregar PDF
          </a>
        </div>
        <div style={{ ...S.hint, marginTop: 8 }}>
          O PDF foi pensado para onboarding multiempresa. A versao nova passa a cobrir projetos, leads, tarefas e tickets, mantendo configuracao independente por entidade.
        </div>
      </div>

      {validationError && (
        <PanelState
          compact
          tone="error"
          title="Falha na validacao Odoo"
          description={validationError}
        />
      )}

      {validation && (
        <div style={{ display: "grid", gap: 10 }}>
          <PanelState
            compact
            tone={validation.ready ? "success" : validation.summary.error > 0 ? "error" : "warning"}
            title={validation.ready ? "Layout pronto para uso" : "Layout requer ajustes no Odoo Studio"}
            description={`Modelo ${validation.model}. ${validation.summary.ok} ok, ${validation.summary.warning} aviso(s), ${validation.summary.error} erro(s). Alvo validado: ${validation.target === "lead" ? "Lead" : validation.target === "task" ? "Tarefa" : validation.target === "ticket" ? "Ticket" : "Projeto"}.`}
          />

          <div style={S.referenceCard}>
            <div style={S.fieldLabel}>Checklist da validacao</div>
            <div style={{ display: "grid", gap: 8 }}>
              {validation.checks.map((check) => {
                const toneColor = check.status === "ok" ? "var(--iccc-pill-active-bg)" : check.status === "warning" ? "#d97706" : "#ef4444";
                const toneBg = check.status === "ok" ? "rgba(16, 185, 129, 0.08)" : check.status === "warning" ? "rgba(217, 119, 6, 0.08)" : "rgba(239, 68, 68, 0.08)";
                return (
                  <div
                    key={check.key}
                    style={{
                      borderRadius: 12,
                      border: `1px solid ${toneColor}`,
                      background: toneBg,
                      padding: 10,
                      display: "grid",
                      gap: 4,
                    }}
                  >
                    <div style={{ display: "flex", justifyContent: "space-between", gap: 8, alignItems: "center", flexWrap: "wrap" }}>
                      <div style={{ fontSize: 12, fontWeight: 800, color: "var(--iccc-text)" }}>{check.label}</div>
                      <div style={{ fontSize: 10, fontWeight: 800, textTransform: "uppercase", color: toneColor }}>
                        {check.kind === "field" ? "Campo" : "Aba"} · {check.status}
                      </div>
                    </div>
                    <div style={S.hint}>
                      Configurado: <b>{check.configuredName || "Por definir"}</b>
                      {check.actualType ? <> · Tipo real: <b>{check.actualType}</b></> : null}
                    </div>
                    <div style={S.hint}>{check.message}</div>
                    {check.expectedTypes?.length ? (
                      <div style={S.hint}>
                        Tipos aceites: <b>{check.expectedTypes.join(", ")}</b>
                        {check.recommendedType ? <> · Recomendado: <b>{check.recommendedType}</b></> : null}
                      </div>
                    ) : null}
                    {typeof check.presentInFormView === "boolean" ? (
                      <div style={S.hint}>
                        Vista form: <b>{check.presentInFormView ? "Campo presente" : "Campo nao visivel"}</b>
                      </div>
                    ) : null}
                    {check.details ? <div style={S.hint}>{check.details}</div> : null}
                  </div>
                );
              })}
            </div>
          </div>

          <div style={S.referenceCard}>
            <div style={S.fieldLabel}>Vista form detetada</div>
            <div style={{ display: "grid", gap: 6 }}>
              <div style={S.hint}>
                Leitura da vista: <b>{validation.formView?.available ? "OK" : "Nao confirmada"}</b>
              </div>
              <div style={S.hint}>
                Abas encontradas: <b>{validation.formView?.tabTitles?.length ? validation.formView?.tabTitles.join(" | ") : "Nenhuma identificada"}</b>
              </div>
              {validation.formView?.error ? <div style={S.hint}>{validation.formView.error}</div> : null}
            </div>
          </div>
        </div>
      )}

      <PanelState
        compact
        tone="success"
        title="Estado desta fase do CRM2"
        description="Os settings, a validacao Odoo Studio e o DialogApp ja suportam projetos, leads, tarefas e tickets com configuracao independente por entidade, incluindo fallback automatico para descricao."
      />
    </div>
  );
}

function ProtectionSettings() {
  const [data, setData] = useState<string[][]>([]);
  const [mapping, setMapping] = useState<Record<string, string>>({});
  const [isMapping, setIsMapping] = useState(false);
  const [status, setStatus] = useState("");

  async function onFileDrop(e: React.ChangeEvent<HTMLInputElement>) {
    const file = e.target.files?.[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = async () => {
      const text = String(reader.result || "");
      const rows = text.split("\n").map(r => r.split(",").map(c => c.trim()));
      const headers = rows[0] || [];
      setData(rows);

      setIsMapping(true);
      const { mapHeadersAi } = await import("../modules/crm/excelProvider");
      const m = await mapHeadersAi(headers);
      setMapping(m);
      setIsMapping(false);
    };
    reader.readAsText(file);
  }

  async function onSave() {
    if (data.length < 2) return;
    setStatus("A guardar...");
    const { saveProjects } = await import("../modules/crm/excelProvider");

    const headers = data[0];
    const projects = data.slice(1).map(row => {
      const p: any = { refArticles: [] };
      row.forEach((val, idx) => {
        const key = mapping[headers[idx]];
        if (key) {
          if (key === "refArticles") p.refArticles.push(val);
          else p[key] = val;
        }
      });
      return p;
    }).filter(p => p.projectName);

    await saveProjects(projects);
    setStatus("✓ Tabela de proteção atualizada localmente.");
    setTimeout(() => setStatus(""), 3000);
  }

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <div style={S.hint}>
        Carrega o teu ficheiro de proteção (CSV). A IA mapeia as colunas automaticamente.
        Os dados ficam guardados apenas no teu browser (**IndexedDB**).
      </div>

      <div style={{
        border: "2px dashed var(--iccc-card-border)",
        borderRadius: 12,
        padding: 20,
        cursor: "pointer"
      }}>
        <input type="file" accept=".csv" onChange={onFileDrop} style={{ display: "none" }} id="moat-upload" />
        <label htmlFor="moat-upload" style={{ cursor: "pointer" }}>
          <Icons.Upload size={24} style={{ marginBottom: 8, opacity: 0.5 }} />
          <div style={{ fontSize: 13, fontWeight: 700 }}>Arrasta ou clica para carregar Excel/CSV</div>
          <div style={{ fontSize: 11, opacity: 0.7 }}>Bypass IT: Local-First Storage</div>
        </label>
      </div>

      {data.length > 0 && (
        <div style={{ padding: 12, background: "rgba(0,0,0,0.02)", borderRadius: 12, border: "1px solid var(--iccc-card-border)" }}>
          <div style={{ display: "flex", alignItems: "center", gap: 8, marginBottom: 10 }}>
            <Icons.Sparkles size={14} color="#2563eb" />
            <div style={{ fontSize: 11, fontWeight: 800, textTransform: "uppercase" }}>Mapeamento IA</div>
          </div>
          {isMapping ? <div style={{ fontSize: 11 }}>A analisar colunas...</div> : (
            <div style={{ display: "grid", gap: 4 }}>
              {Object.entries(mapping).map(([h, internal]) => (
                <div key={h} style={{ display: "flex", justifyContent: "space-between", fontSize: 11 }}>
                  <span style={{ opacity: 0.7 }}>{h}</span>
                  <Icons.ArrowRight size={10} style={{ margin: "0 6px" }} />
                  <span style={{ fontWeight: 700, color: "#2563eb" }}>{internal}</span>
                </div>
              ))}
            </div>
          )}
          <button style={{ ...S.btn, width: "100%", marginTop: 12 }} onClick={onSave}>
            Confirmar e Sincronizar Localmente
          </button>
        </div>
      )}

      {status && <div style={S.okBox}>{status}</div>}
    </div>
  );
}

function ConnectionSettings({ model, setModel, setStatus, availableModels, fetchingModels, refreshModels }: {
  model: CockpitSettingsV1,
  setModel: (s: CockpitSettingsV1) => void,
  setStatus: (s: StatusValue) => void,
  availableModels: { openai: string[]; gemini: string[] },
  fetchingModels: boolean,
  refreshModels: () => Promise<void>
}) {
  const { granularStatus, granularStatusDetails, checkConnectivity, login } = useCockpit();
  const [isTesting, setIsTesting] = useState(false);

  const handleTest = async () => {
    setIsTesting(true);
    setStatus({
      tone: "loading",
      title: "A testar ligações",
      description: "Estamos a validar o acesso ao Odoo e aos fornecedores de IA.",
    });
    try {
      // 1. Odoo Login/Session test
      await login({
        url: model.odooUrl,
        db: model.odooDb,
        login: model.odooLogin,
        password: model.odooPassword
      });

      // 2. Complete check (Odoo Ping + AI Selftests)
      const customModels = {
        openaiModelFast: model.openaiModelFast,
        openaiModelQuality: model.openaiModelQuality,
        geminiModel: model.geminiModel,
        openaiApiKey: model.openaiApiKey,
        geminiApiKey: model.geminiApiKey,
      };
      await checkConnectivity(customModels);
      setStatus("Ligações testadas com sucesso.");
    } catch (e: any) {
      console.error("[Settings] Test failed:", e);
      if (typeof setStatus === "function") {
        setStatus(`Erro no teste: ${e.message || String(e)}`);
      }
    } finally {
      setIsTesting(false);
    }
  };

  const StatusDot = ({ ok }: { ok: boolean | null }) => {
    const color = ok === null ? "#ccc" : ok ? "#36b37e" : "#ff5630";
    const label = ok === null ? "Por testar" : ok ? "Ligação Ativa" : "Falha na Ligação";
    return (
      <div style={{
        display: "flex",
        alignItems: "center",
        gap: 4,
        fontSize: 10,
        fontWeight: 700,
        color
      }}>
        <div style={{
          width: 6,
          height: 6,
          borderRadius: "50%",
          background: color
        }} />
        {label}
      </div>
    );
  };

  return (
    <div style={{ display: "grid", gap: 12 }}>
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={S.fieldLabel}>Odoo Integration</div>
        <StatusDot ok={granularStatus.odoo} />
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="URL da Instância">
          <input
            style={S.input}
            placeholder="https://suaempresa.odoo.com"
            value={model.odooUrl || ""}
            onChange={e => setModel({ ...model, odooUrl: e.target.value })}
          />
        </Field>
        <Field label="Base de Dados">
          <input
            style={S.input}
            placeholder="db_name"
            value={model.odooDb || ""}
            onChange={e => setModel({ ...model, odooDb: e.target.value })}
          />
        </Field>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Utilizador (Login)">
          <input
            style={S.input}
            placeholder="pedro@empresa.com"
            value={model.odooLogin || ""}
            onChange={e => setModel({ ...model, odooLogin: e.target.value })}
          />
        </Field>
        <Field label="Password / Token">
          <input
            type="password"
            style={S.input}
            placeholder="••••••••"
            value={model.odooPassword || ""}
            onChange={e => setModel({ ...model, odooPassword: e.target.value })}
          />
        </Field>
      </div>

      <hr style={{ border: "none", borderTop: "1px solid var(--iccc-card-border)", margin: "4px 0" }} />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={S.fieldLabel}>InvoiceStudio</div>
        <label style={{ display: "inline-flex", alignItems: "center", gap: 6, fontSize: 11, fontWeight: 700, color: "var(--iccc-text-muted)" }}>
          <input
            type="checkbox"
            checked={model.invoiceStudio.enabled === true}
            onChange={e => setModel({
              ...model,
              invoiceStudio: {
                ...model.invoiceStudio,
                enabled: e.target.checked,
              }
            })}
          />
          Ativo
        </label>
      </div>
      <div style={{ fontSize: 11, color: "var(--iccc-text-muted)", marginTop: -6 }}>
        Integração isolada para envio de anexos da aba FILES para processamento no InvoiceStudio.
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="URL do Backend">
          <input
            style={S.input}
            placeholder="https://invoice-studio-backend.onrender.com"
            value={model.invoiceStudio.baseUrl || ""}
            onChange={e => setModel({
              ...model,
              invoiceStudio: {
                ...model.invoiceStudio,
                baseUrl: e.target.value,
              }
            })}
          />
        </Field>
        <Field label="Projeto / Workspace">
          <input
            style={S.input}
            placeholder="Dvtkb"
            value={model.invoiceStudio.project || ""}
            onChange={e => setModel({
              ...model,
              invoiceStudio: {
                ...model.invoiceStudio,
                project: e.target.value,
              }
            })}
          />
        </Field>
      </div>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Email técnico">
          <input
            style={S.input}
            placeholder="integration@example.com"
            value={model.invoiceStudio.email || ""}
            onChange={e => setModel({
              ...model,
              invoiceStudio: {
                ...model.invoiceStudio,
                email: e.target.value,
              }
            })}
          />
        </Field>
        <Field label="Password">
          <input
            type="password"
            style={S.input}
            placeholder="••••••••"
            value={model.invoiceStudio.password || ""}
            onChange={e => setModel({
              ...model,
              invoiceStudio: {
                ...model.invoiceStudio,
                password: e.target.value,
              }
            })}
          />
        </Field>
      </div>

      <hr style={{ border: "none", borderTop: "1px solid var(--iccc-card-border)", margin: "4px 0" }} />
      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={S.fieldLabel}>AI Intelligence (OpenAI)</div>
        <StatusDot ok={granularStatus.openai} />
      </div>
      {granularStatusDetails.openai && (
        <div style={{ fontSize: 11, color: "var(--iccc-status-error)", marginBottom: 8 }}>
          {granularStatusDetails.openai}
        </div>
      )}
      <Field label="OpenAI API Key (Opcional se definida no server)">
        <input
          type="password"
          style={S.input}
          placeholder="Introduza a sua API Key..."
          value={model.openaiApiKey || ""}
          onChange={e => setModel({ ...model, openaiApiKey: e.target.value })}
        />
      </Field>

      <div style={{ display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10 }}>
        <Field label="Modelo Rápido (OpenAI)">
          <select
            style={S.select}
            value={model.openaiModelFast || ""}
            onChange={e => setModel({ ...model, openaiModelFast: e.target.value })}
          >
            {availableModels.openai.length > 0 ? (
              availableModels.openai.map(m => <option key={m} value={m}>{m}</option>)
            ) : null}
            <option value="">Usar padrão do servidor</option>
          </select>
        </Field>
        <Field label="Modelo Qualidade (OpenAI)">
          <select
            style={S.select}
            value={model.openaiModelQuality || ""}
            onChange={e => setModel({ ...model, openaiModelQuality: e.target.value })}
          >
            {availableModels.openai.length > 0 ? (
              availableModels.openai.map(m => <option key={m} value={m}>{m}</option>)
            ) : null}
            <option value="">Usar padrão do servidor</option>
          </select>
        </Field>
      </div>

      <hr style={{ border: "none", borderTop: "1px solid var(--iccc-card-border)", margin: "4px 0" }} />

      <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center" }}>
        <div style={{ display: "flex", alignItems: "center", gap: 8 }}>
          <div style={S.fieldLabel}>AI Intelligence (Gemini)</div>
          <StatusDot ok={granularStatus.gemini} />
        </div>
        <button
          style={{ ...S.btnGhost, padding: "2px 8px", fontSize: 9 }}
          onClick={refreshModels}
          disabled={fetchingModels}
        >
          {fetchingModels ? "A procurar..." : "Localizar Modelos"}
        </button>
      </div>
      {granularStatusDetails.geminiDetails && (
        <div style={{
          fontSize: 10,
          color: "var(--iccc-text-muted)",
          marginBottom: 8,
          padding: 8,
          background: "rgba(0,0,0,0.02)",
          borderRadius: 8,
          border: "1px solid var(--iccc-card-border)"
        }}>
          <div><b>Model Request:</b> {granularStatusDetails.geminiDetails.requested}</div>
          <div><b>Effective Model:</b> {granularStatusDetails.geminiDetails.effective}</div>
          {granularStatusDetails.geminiDetails.requested !== granularStatusDetails.geminiDetails.effective && (
            <div style={{ color: "#f59e0b", marginTop: 4 }}>
              <Icons.AlertTriangle size={10} style={{ marginRight: 4 }} />
              Fallback ativado (modelo indisponível ou inválido).
            </div>
          )}
        </div>
      )}
      <Field label="Gemini API Key">
        <input
          type="password"
          style={S.input}
          placeholder="Introduza a sua API Key..."
          value={model.geminiApiKey || ""}
          onChange={e => setModel({ ...model, geminiApiKey: e.target.value })}
        />
      </Field>

      <Field label="Modelo Gemini (3.1 Flash/Pro)">
        <select
          style={S.select}
          value={model.geminiModel || ""}
          onChange={e => setModel({ ...model, geminiModel: e.target.value })}
        >
          <option value="">Usar padrão do servidor</option>
          {availableModels.gemini.length > 0 ? (
            availableModels.gemini.map(m => (
              <option key={m} value={m}>{m}</option>
            ))
          ) : null}
        </select>
      </Field>

      <button
        style={{ ...S.btn, marginTop: 10, background: "#0f172a", width: "100%" }}
        onClick={handleTest}
        disabled={isTesting}
      >
        {isTesting ? "A Testar..." : "Testar Ligações"}
      </button>

      <div style={S.hint}>
        Nota: Mantém o Odoo aberto no browser para acesso direto sem login. Clique em "Testar Ligações" para validar o acesso ao RPC e health checks do Gemini.
      </div>
    </div>
  );
}
