import React, { useEffect, useMemo, useState } from "react";
import * as Icons from "@/ui/icons";
import { HelpHint } from "@/ui/HelpHint";
import {
  clearSignatureImageDataUrl,
  getSettings,
  getSignatureImageDataUrl,
  saveSettings,
  setSignatureImageDataUrl,
  type AiAutoLabelId,
  type AiCustomTone,
  type AiTextShortcut,
  type AppLocale,
  type CockpitSettingsV1,
  type ContactAlias,
  type ResponsePreset,
} from "@/settings";
import { requestCockpitHostAction } from "@/office";

type SectionId =
  | "general"
  | "ai-knowledge"
  | "response-presets"
  | "nicknames"
  | "signature"
  | "custom-tones"
  | "auto-labels"
  | "text-shortcuts"
  | "font-preference";

const LOCALE_LABEL: Record<AppLocale, string> = {
  "pt-PT": "Português (Portugal)",
  "es-ES": "Espanhol (Espanha)",
  "en-GB": "Inglês (UK)",
  "it-IT": "Italiano (IT)",
  "de-DE": "Alemão (DE)",
};

const LANG_OPTIONS: Array<{ value: CockpitSettingsV1["readingLanguage"]; label: string }> = [
  { value: "auto", label: "Auto" },
  { value: "pt-PT", label: "Português (PT)" },
  { value: "es-ES", label: "Espanhol (ES)" },
  { value: "en-GB", label: "Inglês (UK)" },
  { value: "it-IT", label: "Italiano (IT)" },
  { value: "de-DE", label: "Alemão (DE)" },
];

const TONE_OPTIONS: Array<{ value: CockpitSettingsV1["tone"]; label: string }> = [
  { value: "neutro", label: "Neutro" },
  { value: "curto", label: "Curto" },
  { value: "direto", label: "Direto" },
  { value: "simpático", label: "Simpático" },
];

const LENGTH_OPTIONS: Array<{ value: CockpitSettingsV1["length"]; label: string }> = [
  { value: "xs", label: "Extra curta" },
  { value: "s", label: "Curta" },
  { value: "m", label: "Média" },
  { value: "l", label: "Longa" },
];

const FONT_FAMILIES = ["Segoe UI", "Aptos", "Calibri", "Arial", "Georgia", "Tahoma", "Verdana"];

const AUTO_LABEL_META: Array<{ id: AiAutoLabelId; label: string; description: string; tone: string }> = [
  { id: "to_respond", label: "TO RESPOND", description: "Requer resposta ou ação.", tone: "#F5DCC7" },
  { id: "meeting", label: "MEETING", description: "Convite ou agendamento.", tone: "#F6D3D8" },
  { id: "fyi", label: "FYI", description: "Informativo; sem resposta necessária.", tone: "#D9EEF9" },
  { id: "notification", label: "NOTIFICATION", description: "Comunicação pessoal ou social.", tone: "#D7F2F0" },
  { id: "internal_update", label: "INTERNAL UPDATE", description: "Atualização interna.", tone: "#E4E1FA" },
  { id: "awaiting_reply", label: "AWAITING REPLY", description: "Estamos à espera de resposta.", tone: "#F9EBC9" },
  { id: "marketing", label: "MARKETING", description: "Promoção, campanha ou evento.", tone: "#EEF4D3" },
  { id: "done", label: "DONE", description: "Sem ação pendente do teu lado.", tone: "#D8F0D7" },
];

const MENU: Array<{ id: SectionId; label: string; icon: React.ReactNode }> = [
  { id: "general", label: "General", icon: <Icons.Settings size={15} /> },
  { id: "ai-knowledge", label: "AI knowledge", icon: <Icons.Sparkles size={15} /> },
  { id: "response-presets", label: "MODS", icon: <Icons.Clipboard size={15} /> },
  { id: "nicknames", label: "Nicknames", icon: <Icons.MessageSquare size={15} /> },
  { id: "signature", label: "Signature", icon: <Icons.Edit size={15} /> },
  { id: "custom-tones", label: "Custom tones", icon: <Icons.RefreshCw size={15} /> },
  { id: "auto-labels", label: "Auto label & drafts", icon: <Icons.Star size={15} /> },
  { id: "text-shortcuts", label: "Text shortcuts", icon: <Icons.Clipboard size={15} /> },
  { id: "font-preference", label: "Font preference", icon: <Icons.Files size={15} /> },
];

function uid(prefix: string) {
  return `${prefix}-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`;
}

function SectionHeader({ title, description }: { title: string; description: string }) {
  return (
    <div style={{ display: "grid", gap: 6 }}>
      <div style={{ fontSize: 28, fontWeight: 800, color: "#111827" }}>{title}</div>
      <div style={{ fontSize: 13, lineHeight: 1.45, color: "#6B7280" }}>{description}</div>
    </div>
  );
}

function Field({ label, help, children }: { label: string; help?: string; children: React.ReactNode }) {
  return (
    <div style={{ display: "grid", gap: 6 }}>
      <div style={S.fieldLabelRow}>
        <div style={S.fieldLabel}>{label}</div>
        {help ? <HelpHint text={help} /> : null}
      </div>
      {children}
    </div>
  );
}

function EmptyState({ text }: { text: string }) {
  return <div style={S.emptyState}>{text}</div>;
}

export default function AiSettingsApp() {
  const [section, setSection] = useState<SectionId>("general");
  const [model, setModel] = useState<CockpitSettingsV1 | null>(null);
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<{ tone: "success" | "error"; text: string } | null>(null);
  const [activeLocale, setActiveLocale] = useState<AppLocale>("pt-PT");
  const [sigImgLocal, setSigImgLocal] = useState<Partial<Record<AppLocale, string>>>({});

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const settings = await getSettings();
        if (!alive) return;
        setModel(settings);
        setSigImgLocal({
          "pt-PT": getSignatureImageDataUrl("pt-PT"),
          "es-ES": getSignatureImageDataUrl("es-ES"),
          "en-GB": getSignatureImageDataUrl("en-GB"),
          "it-IT": getSignatureImageDataUrl("it-IT"),
          "de-DE": getSignatureImageDataUrl("de-DE"),
        });
      } catch (error: any) {
        if (!alive) return;
        setStatus({ tone: "error", text: error?.message || "Não foi possível carregar as definições da IA." });
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  const currentSigImage = useMemo(() => {
    if (!model) return "";
    return String(sigImgLocal[activeLocale] || model.signatureImageUrl?.[activeLocale] || "").trim();
  }, [activeLocale, model, sigImgLocal]);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (!closed) window.close();
  }

  async function handleSave() {
    if (!model) return;
    setSaving(true);
    setStatus(null);
    try {
      const next = await saveSettings({
        readingLanguage: model.readingLanguage,
        replyLanguage: model.replyLanguage,
        tone: model.tone,
        length: model.length,
        aiManualOnly: model.aiManualOnly,
        aiKnowledge: model.aiKnowledge,
        responsePresets: model.responsePresets,
        contactAliases: model.contactAliases,
        userRole: model.userRole,
        styleContext: model.styleContext,
        styleExamples: model.styleExamples,
        meetingLinks: model.meetingLinks,
        signatures: model.signatures,
        signaturesHtml: model.signaturesHtml,
        signatureImageUrl: model.signatureImageUrl,
        signatureImageMaxWidth: model.signatureImageMaxWidth,
        aiCustomTones: model.aiCustomTones,
        aiTextShortcuts: model.aiTextShortcuts,
        aiAutoLabel: model.aiAutoLabel,
        aiFontPreference: model.aiFontPreference,
      });
      setModel(next);
      setStatus({ tone: "success", text: "Definições da IA guardadas." });
    } catch (error: any) {
      setStatus({ tone: "error", text: error?.message || "Falha ao guardar as definições da IA." });
    } finally {
      setSaving(false);
    }
  }

  function addAlias() {
    if (!model) return;
    setModel({ ...model, contactAliases: [...model.contactAliases, { id: uid("alias"), name: "", email: "" }] });
  }

  function updateAlias(id: string, patch: Partial<ContactAlias>) {
    if (!model) return;
    setModel({ ...model, contactAliases: model.contactAliases.map((entry) => (entry.id === id ? { ...entry, ...patch } : entry)) });
  }

  function removeAlias(id: string) {
    if (!model) return;
    setModel({ ...model, contactAliases: model.contactAliases.filter((entry) => entry.id !== id) });
  }

  function addResponsePreset() {
    if (!model) return;
    const preset: ResponsePreset = { id: uid("preset"), name: "Novo MOD", prompt: "" };
    setModel({ ...model, responsePresets: [...(model.responsePresets || []), preset] });
  }

  function updateResponsePreset(id: string, patch: Partial<ResponsePreset>) {
    if (!model) return;
    setModel({
      ...model,
      responsePresets: (model.responsePresets || []).map((entry) => (entry.id === id ? { ...entry, ...patch } : entry)),
    });
  }

  function duplicateResponsePreset(id: string) {
    if (!model) return;
    const presets = model.responsePresets || [];
    const preset = presets.find((entry) => entry.id === id);
    if (!preset) return;
    const copy: ResponsePreset = {
      ...preset,
      id: uid("preset"),
      name: `${preset.name || "MOD"} (copia)`,
    };
    const idx = presets.findIndex((entry) => entry.id === id);
    const next = [...presets];
    next.splice(idx + 1, 0, copy);
    setModel({ ...model, responsePresets: next });
  }

  function removeResponsePreset(id: string) {
    if (!model) return;
    setModel({ ...model, responsePresets: (model.responsePresets || []).filter((entry) => entry.id !== id) });
  }

  function moveResponsePreset(id: string, direction: -1 | 1) {
    if (!model) return;
    const presets = [...(model.responsePresets || [])];
    const idx = presets.findIndex((entry) => entry.id === id);
    const target = idx + direction;
    if (idx < 0 || target < 0 || target >= presets.length) return;
    const [entry] = presets.splice(idx, 1);
    presets.splice(target, 0, entry);
    setModel({ ...model, responsePresets: presets });
  }

  function addCustomTone() {
    if (!model) return;
    setModel({ ...model, aiCustomTones: [...model.aiCustomTones, { id: uid("tone"), name: "", instructions: "" }] });
  }

  function updateCustomTone(id: string, patch: Partial<AiCustomTone>) {
    if (!model) return;
    setModel({ ...model, aiCustomTones: model.aiCustomTones.map((entry) => (entry.id === id ? { ...entry, ...patch } : entry)) });
  }

  function removeCustomTone(id: string) {
    if (!model) return;
    setModel({ ...model, aiCustomTones: model.aiCustomTones.filter((entry) => entry.id !== id) });
  }

  function addShortcut() {
    if (!model) return;
    setModel({ ...model, aiTextShortcuts: [...model.aiTextShortcuts, { id: uid("shortcut"), trigger: "", content: "" }] });
  }

  function updateShortcut(id: string, patch: Partial<AiTextShortcut>) {
    if (!model) return;
    setModel({ ...model, aiTextShortcuts: model.aiTextShortcuts.map((entry) => (entry.id === id ? { ...entry, ...patch } : entry)) });
  }

  function removeShortcut(id: string) {
    if (!model) return;
    setModel({ ...model, aiTextShortcuts: model.aiTextShortcuts.filter((entry) => entry.id !== id) });
  }

  async function uploadSignature(loc: AppLocale, file: File) {
    const dataUrl = await new Promise<string>((resolve, reject) => {
      const reader = new FileReader();
      reader.onerror = () => reject(new Error("Falha ao ler a imagem."));
      reader.onload = () => resolve(String(reader.result || ""));
      reader.readAsDataURL(file);
    });
    setSignatureImageDataUrl(loc, dataUrl);
    setSigImgLocal((prev) => ({ ...prev, [loc]: dataUrl }));
  }

  function clearSignatureUpload(loc: AppLocale) {
    clearSignatureImageDataUrl(loc);
    setSigImgLocal((prev) => ({ ...prev, [loc]: "" }));
  }

  if (loading || !model) {
    return (
      <div style={S.loadingRoot}>
        <div style={S.loadingCard}>
          <Icons.RotateCcw size={16} style={{ animation: "spin 1s linear infinite" }} />
          <span>A carregar settings da IA...</span>
        </div>
        <style>{`@keyframes spin { from { transform: rotate(0deg); } to { transform: rotate(360deg); } }`}</style>
      </div>
    );
  }

  return (
    <div style={S.root}>
      <div style={S.window}>
        <div style={S.header}>
          <div>
            <div style={S.headerEyebrow}>Settings</div>
            <div style={S.headerTitle}>AI Settings</div>
          </div>
          <div style={{ display: "flex", gap: 8 }}>
            <button type="button" style={S.ghostBtn} onClick={handleClose}>Fechar</button>
            <button type="button" style={S.primaryBtn} onClick={handleSave} disabled={saving}>
              {saving ? "A guardar..." : "Guardar"}
            </button>
          </div>
        </div>

        <div style={S.body}>
          <aside style={S.sidebar}>
            {MENU.map((item) => (
              <button key={item.id} type="button" style={section === item.id ? S.sideItemOn : S.sideItem} onClick={() => setSection(item.id)}>
                <span style={{ display: "inline-flex", opacity: section === item.id ? 1 : 0.72 }}>{item.icon}</span>
                <span>{item.label}</span>
              </button>
            ))}
            <div style={S.sidebarHelp}>Help</div>
          </aside>

          <section style={S.content}>
            {status ? (
              <div style={{
                borderRadius: 12,
                border: `1px solid ${status.tone === "success" ? "rgba(22, 163, 74, 0.28)" : "rgba(220, 38, 38, 0.24)"}`,
                background: status.tone === "success" ? "rgba(22, 163, 74, 0.08)" : "rgba(220, 38, 38, 0.08)",
                color: status.tone === "success" ? "#166534" : "#991B1B",
                padding: "10px 12px",
                fontSize: 12,
                fontWeight: 700,
              }}>{status.text}</div>
            ) : null}

            {section === "general" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="General" description="Base da experiência da aba IA. Aqui defines idioma, tom, comprimento e se a IA trabalha só por clique." />
                <Field label="Reading language" help="Idioma base para ler, resumir e interpretar emails.">
                  <select style={S.select} value={model.readingLanguage} onChange={(e) => setModel({ ...model, readingLanguage: e.target.value as CockpitSettingsV1["readingLanguage"] })}>
                    {LANG_OPTIONS.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
                  </select>
                </Field>
                <Field label="Reply language" help="Idioma de saída das respostas geradas pela IA.">
                  <select style={S.select} value={model.replyLanguage} onChange={(e) => setModel({ ...model, replyLanguage: e.target.value as CockpitSettingsV1["replyLanguage"] })}>
                    {LANG_OPTIONS.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
                  </select>
                </Field>
                <div style={S.grid2}>
                  <Field label="Default tone" help="Tom base das respostas e drafts.">
                    <select style={S.select} value={model.tone} onChange={(e) => setModel({ ...model, tone: e.target.value as CockpitSettingsV1["tone"] })}>
                      {TONE_OPTIONS.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
                    </select>
                  </Field>
                  <Field label="Default length" help="Comprimento sugerido das respostas geradas.">
                    <select style={S.select} value={model.length} onChange={(e) => setModel({ ...model, length: e.target.value as CockpitSettingsV1["length"] })}>
                      {LENGTH_OPTIONS.map((opt) => <option key={opt.value} value={opt.value}>{opt.label}</option>)}
                    </select>
                  </Field>
                </div>
                <label style={S.toggleRow}>
                  <input type="checkbox" checked={model.aiManualOnly !== false} onChange={(e) => setModel({ ...model, aiManualOnly: e.target.checked })} />
                  <div>
                    <div style={S.toggleTitle}>Manual only</div>
                    <div style={S.hint}>Quando ativo, a aba IA só chama modelos por clique explícito.</div>
                  </div>
                </label>
              </div>
            ) : null}

            {section === "ai-knowledge" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="AI knowledge" description="Factos, regras e notas persistentes que a IA deve respeitar ao escrever e resumir." />
                <Field label="Knowledge base" help="Uma linha por regra ou dado importante.">
                  <textarea style={{ ...S.textarea, minHeight: 220 }} value={(model.aiKnowledge || []).join("\n")} onChange={(e) => setModel({ ...model, aiKnowledge: e.target.value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean) })} placeholder={"NIF: ...\nIBAN: ...\nResponder sempre em tom profissional."} />
                </Field>
                <Field label="Writing profile" help="Contexto geral do teu papel e do teu estilo.">
                  <input style={S.input} value={model.userRole || ""} onChange={(e) => setModel({ ...model, userRole: e.target.value })} placeholder="Ex.: Gestor comercial na Divitek" />
                </Field>
                <Field label="Style context" help="Como costumas escrever e o que queres que a IA imite.">
                  <textarea style={{ ...S.textarea, minHeight: 110 }} value={model.styleContext || ""} onChange={(e) => setModel({ ...model, styleContext: e.target.value })} placeholder="Ex.: Direto, claro, sem excesso de formalismo." />
                </Field>
                <Field label="Writing examples" help="Exemplos reais teus para a IA captar ritmo e estrutura.">
                  <textarea style={{ ...S.textarea, minHeight: 180 }} value={model.styleExamples || ""} onChange={(e) => setModel({ ...model, styleExamples: e.target.value })} placeholder="Cola aqui 2 ou 3 exemplos de emails teus." />
                </Field>
              </div>
            ) : null}

            {section === "response-presets" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="MODS / Response Presets" description="Instrucoes reutilizaveis para gerar respostas com o contexto real do email. Esta e a fonte oficial do menu MODS do cockpit." />
                <div style={S.toolbarRow}>
                  <button type="button" style={S.ghostBtn} onClick={addResponsePreset}><Icons.Plus size={14} /> Adicionar MOD</button>
                </div>
                <div style={S.listStack}>
                  {(model.responsePresets || []).map((preset, index) => (
                    <div key={preset.id} style={S.blockCard}>
                      <div style={S.inlineCard}>
                        <input
                          style={{ ...S.input, flex: 1 }}
                          value={preset.name}
                          onChange={(e) => updateResponsePreset(preset.id, { name: e.target.value })}
                          placeholder="Nome do MOD"
                        />
                        <button type="button" style={S.ghostBtn} onClick={() => moveResponsePreset(preset.id, -1)} disabled={index === 0}>Subir</button>
                        <button type="button" style={S.ghostBtn} onClick={() => moveResponsePreset(preset.id, 1)} disabled={index === (model.responsePresets || []).length - 1}>Descer</button>
                        <button type="button" style={S.ghostBtn} onClick={() => duplicateResponsePreset(preset.id)}>Duplicar</button>
                        <button type="button" style={S.iconBtnDanger} onClick={() => removeResponsePreset(preset.id)}><Icons.Trash size={14} /></button>
                      </div>
                      <textarea
                        style={{ ...S.textarea, minHeight: 120 }}
                        value={preset.prompt}
                        onChange={(e) => updateResponsePreset(preset.id, { prompt: e.target.value })}
                        placeholder="Instrucao para a IA. Ex.: Agradece o contacto e pede o NIF de faturacao."
                      />
                      <div style={S.hint}>O MOD e usado como instrucao explicita; a IA continua a gerar com base no email, contexto e pipeline atual.</div>
                    </div>
                  ))}
                  {(model.responsePresets || []).length === 0 ? <EmptyState text="Sem MODS configurados." /> : null}
                </div>
              </div>
            ) : null}

            {section === "nicknames" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Nicknames" description="Mapa rápido de nomes curtos, fábricas e contactos recorrentes." />
                <div style={S.toolbarRow}>
                  <button type="button" style={S.ghostBtn} onClick={addAlias}><Icons.Plus size={14} /> Adicionar</button>
                </div>
                <div style={S.listStack}>
                  {model.contactAliases.map((alias) => (
                    <div key={alias.id} style={S.inlineCard}>
                      <input style={{ ...S.input, flex: 1.2 }} value={alias.name} onChange={(e) => updateAlias(alias.id, { name: e.target.value })} placeholder="Nome curto" />
                      <input style={{ ...S.input, flex: 1.8 }} value={alias.email} onChange={(e) => updateAlias(alias.id, { email: e.target.value })} placeholder="email@dominio.pt" />
                      <button type="button" style={S.iconBtnDanger} onClick={() => removeAlias(alias.id)}><Icons.Trash size={14} /></button>
                    </div>
                  ))}
                  {model.contactAliases.length === 0 ? <EmptyState text="Sem nicknames definidos." /> : null}
                </div>
              </div>
            ) : null}

            {section === "signature" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Signature" description="Assinatura da IA por idioma. Suporta upload, URL, HTML e fallback em texto." />
                <div style={S.localeRow}>
                  {(Object.keys(LOCALE_LABEL) as AppLocale[]).map((loc) => (
                    <button key={loc} type="button" style={activeLocale === loc ? S.localePillOn : S.localePill} onClick={() => setActiveLocale(loc)}>{loc}</button>
                  ))}
                </div>
                <div style={S.signatureCard}>
                  <div style={S.grid2}>
                    <Field label="Upload image">
                      <input type="file" accept="image/*" onChange={(e) => { const file = e.target.files?.[0]; if (file) void uploadSignature(activeLocale, file); (e.target as HTMLInputElement).value = ""; }} />
                    </Field>
                    <Field label="Image max width (px)">
                      <input type="number" min={120} max={900} style={S.input} value={String(model.signatureImageMaxWidth?.[activeLocale] ?? 260)} onChange={(e) => setModel({ ...model, signatureImageMaxWidth: { ...(model.signatureImageMaxWidth || {}), [activeLocale]: Math.max(120, Math.min(900, Number(e.target.value || 260))) } })} />
                    </Field>
                  </div>
                  <Field label="Image URL">
                    <input style={S.input} value={String(model.signatureImageUrl?.[activeLocale] || "")} onChange={(e) => setModel({ ...model, signatureImageUrl: { ...(model.signatureImageUrl || {}), [activeLocale]: e.target.value } })} placeholder="https://..." />
                  </Field>
                  {currentSigImage ? (
                    <div style={S.previewWrap}>
                      <img src={currentSigImage} alt="" style={{ display: "block", maxWidth: Number(model.signatureImageMaxWidth?.[activeLocale] ?? 260), height: "auto" }} />
                      <button type="button" style={S.ghostBtn} onClick={() => clearSignatureUpload(activeLocale)}>Remover upload local</button>
                    </div>
                  ) : null}
                  <Field label="HTML signature">
                    <textarea style={{ ...S.textarea, minHeight: 150 }} value={String(model.signaturesHtml?.[activeLocale] || "")} onChange={(e) => setModel({ ...model, signaturesHtml: { ...(model.signaturesHtml || {}), [activeLocale]: e.target.value } })} placeholder="<div>Com os melhores cumprimentos...</div>" />
                  </Field>
                  <Field label="Text fallback">
                    <textarea style={{ ...S.textarea, minHeight: 110 }} value={String(model.signatures?.[activeLocale] || "")} onChange={(e) => setModel({ ...model, signatures: { ...(model.signatures || {}), [activeLocale]: e.target.value } })} placeholder={"Com os melhores cumprimentos,\nNome\nEmpresa"} />
                  </Field>
                </div>
              </div>
            ) : null}

            {section === "custom-tones" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Custom tones" description="Tons próprios da tua empresa ou da tua forma de responder." />
                <div style={S.toolbarRow}>
                  <button type="button" style={S.ghostBtn} onClick={addCustomTone}><Icons.Plus size={14} /> Adicionar</button>
                </div>
                <div style={S.listStack}>
                  {model.aiCustomTones.map((tone) => (
                    <div key={tone.id} style={S.blockCard}>
                      <div style={S.inlineCard}>
                        <input style={{ ...S.input, flex: 1 }} value={tone.name} onChange={(e) => updateCustomTone(tone.id, { name: e.target.value })} placeholder="Nome do tom" />
                        <button type="button" style={S.iconBtnDanger} onClick={() => removeCustomTone(tone.id)}><Icons.Trash size={14} /></button>
                      </div>
                      <textarea style={{ ...S.textarea, minHeight: 110 }} value={tone.instructions} onChange={(e) => updateCustomTone(tone.id, { instructions: e.target.value })} placeholder="Ex.: Formal, objetivo, sem excessos, sempre com próximo passo claro." />
                    </div>
                  ))}
                  {model.aiCustomTones.length === 0 ? <EmptyState text="Sem tons personalizados." /> : null}
                </div>
              </div>
            ) : null}

            {section === "auto-labels" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Auto label & drafts" description="Estrutura preparada para classificação automática por IA e drafts contextuais, no mesmo espírito do MailMaestro." />
                <label style={S.toggleRow}>
                  <input type="checkbox" checked={model.aiAutoLabel.enabled} onChange={(e) => setModel({ ...model, aiAutoLabel: { ...model.aiAutoLabel, enabled: e.target.checked } })} />
                  <div>
                    <div style={S.toggleTitle}>Auto labels</div>
                    <div style={S.hint}>Preparação da classificação automática da aba IA.</div>
                  </div>
                </label>
                <label style={S.toggleRow}>
                  <input type="checkbox" checked={model.aiAutoLabel.autoDraftEnabled} onChange={(e) => setModel({ ...model, aiAutoLabel: { ...model.aiAutoLabel, autoDraftEnabled: e.target.checked } })} />
                  <div>
                    <div style={S.toggleTitle}>Auto draft</div>
                    <div style={S.hint}>Permite drafts automáticos quando a classificação o justificar.</div>
                  </div>
                </label>
                <div style={S.listStack}>
                  {AUTO_LABEL_META.map((entry) => (
                    <label key={entry.id} style={S.autoLabelRow}>
                      <input type="checkbox" checked={Boolean(model.aiAutoLabel.labels[entry.id])} onChange={(e) => setModel({ ...model, aiAutoLabel: { ...model.aiAutoLabel, labels: { ...model.aiAutoLabel.labels, [entry.id]: e.target.checked } } })} />
                      <div style={S.autoLabelBody}>
                        <span style={{ ...S.autoLabelChip, background: entry.tone }}>{entry.label}</span>
                        <span style={S.hint}>{entry.description}</span>
                      </div>
                    </label>
                  ))}
                </div>
              </div>
            ) : null}

            {section === "text-shortcuts" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Text shortcuts" description="Snippets rápidos para acelerar prompts, respostas e drafts na aba IA." />
                <div style={S.toolbarRow}>
                  <button type="button" style={S.ghostBtn} onClick={addShortcut}><Icons.Plus size={14} /> Adicionar</button>
                </div>
                <div style={S.listStack}>
                  {model.aiTextShortcuts.map((shortcut) => (
                    <div key={shortcut.id} style={S.blockCard}>
                      <div style={S.inlineCard}>
                        <input style={{ ...S.input, flex: 0.8 }} value={shortcut.trigger} onChange={(e) => updateShortcut(shortcut.id, { trigger: e.target.value })} placeholder="!morada" />
                        <button type="button" style={S.iconBtnDanger} onClick={() => removeShortcut(shortcut.id)}><Icons.Trash size={14} /></button>
                      </div>
                      <textarea style={{ ...S.textarea, minHeight: 100 }} value={shortcut.content} onChange={(e) => updateShortcut(shortcut.id, { content: e.target.value })} placeholder="Texto a inserir quando usares o atalho." />
                    </div>
                  ))}
                  {model.aiTextShortcuts.length === 0 ? <EmptyState text="Sem atalhos de texto." /> : null}
                </div>
              </div>
            ) : null}

            {section === "font-preference" ? (
              <div style={S.sectionStack}>
                <SectionHeader title="Font preference" description="Preferência visual para drafts e saídas futuras da IA." />
                <div style={S.grid2}>
                  <Field label="Font family">
                    <select style={S.select} value={model.aiFontPreference.family} onChange={(e) => setModel({ ...model, aiFontPreference: { ...model.aiFontPreference, family: e.target.value } })}>
                      {FONT_FAMILIES.map((family) => <option key={family} value={family}>{family}</option>)}
                    </select>
                  </Field>
                  <Field label="Font size">
                    <input type="number" min={9} max={20} style={S.input} value={String(model.aiFontPreference.size)} onChange={(e) => setModel({ ...model, aiFontPreference: { ...model.aiFontPreference, size: Math.max(9, Math.min(20, Number(e.target.value || 12))) } })} />
                  </Field>
                </div>
                <Field label="Font color">
                  <div style={{ display: "flex", gap: 10, alignItems: "center" }}>
                    <input type="color" value={model.aiFontPreference.color} onChange={(e) => setModel({ ...model, aiFontPreference: { ...model.aiFontPreference, color: e.target.value } })} />
                    <input style={S.input} value={model.aiFontPreference.color} onChange={(e) => setModel({ ...model, aiFontPreference: { ...model.aiFontPreference, color: e.target.value } })} />
                  </div>
                </Field>
                <div style={S.fontPreview}>
                  <div style={{ fontFamily: model.aiFontPreference.family, fontSize: `${model.aiFontPreference.size}px`, color: model.aiFontPreference.color }}>
                    Exemplo de saída da IA com a preferência de fonte configurada.
                  </div>
                </div>
              </div>
            ) : null}
          </section>
        </div>
      </div>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  root: { height: "100vh", background: "#F7F8FC", color: "#111827", fontFamily: "\"Segoe UI\", system-ui, sans-serif", padding: 16, overflow: "hidden" },
  window: { maxWidth: 1160, margin: "0 auto", height: "calc(100vh - 32px)", borderRadius: 22, background: "#FFFFFF", border: "1px solid #E5E7EB", boxShadow: "0 16px 40px rgba(15, 23, 42, 0.08)", display: "grid", gridTemplateRows: "auto minmax(0, 1fr)", overflow: "hidden" },
  header: { display: "flex", alignItems: "center", justifyContent: "space-between", gap: 12, padding: "18px 22px", borderBottom: "1px solid #EEF2F7" },
  headerEyebrow: { fontSize: 13, fontWeight: 700, color: "#6B7280" },
  headerTitle: { fontSize: 24, fontWeight: 800, color: "#111827" },
  body: { display: "grid", gridTemplateColumns: "280px 1fr", height: "100%", minHeight: 0, overflow: "hidden" },
  sidebar: { display: "grid", alignContent: "start", gap: 4, padding: 14, borderRight: "1px solid #EEF2F7", background: "#FBFCFF", height: "100%", minHeight: 0, overflowY: "auto" },
  sideItem: { display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left", border: "none", background: "transparent", borderRadius: 12, padding: "10px 12px", color: "#4B5563", fontSize: 14, fontWeight: 600, cursor: "pointer" },
  sideItemOn: { display: "flex", alignItems: "center", gap: 10, width: "100%", textAlign: "left", border: "1px solid #DDE3F4", background: "#EEF2FF", borderRadius: 12, padding: "10px 12px", color: "#1D4ED8", fontSize: 14, fontWeight: 700, cursor: "pointer" },
  sidebarHelp: { marginTop: 22, padding: "10px 12px", fontSize: 13, fontWeight: 700, color: "#6B7280" },
  content: { padding: 22, height: "100%", overflowY: "auto", minHeight: 0, display: "grid", alignContent: "start", gap: 16 },
  sectionStack: { display: "grid", gap: 18 },
  fieldLabelRow: { display: "flex", alignItems: "center", gap: 6 },
  fieldLabel: { fontSize: 12, fontWeight: 800, letterSpacing: "0.02em", color: "#374151" },
  hint: { fontSize: 12, lineHeight: 1.45, color: "#6B7280" },
  input: { width: "100%", borderRadius: 12, border: "1px solid #D7DDEA", background: "#FFFFFF", color: "#111827", padding: "10px 12px", fontSize: 13, outline: "none" },
  select: { width: "100%", borderRadius: 12, border: "1px solid #D7DDEA", background: "#FFFFFF", color: "#111827", padding: "10px 12px", fontSize: 13, outline: "none" },
  textarea: { width: "100%", borderRadius: 14, border: "1px solid #D7DDEA", background: "#FFFFFF", color: "#111827", padding: 12, fontSize: 13, lineHeight: 1.45, outline: "none", resize: "vertical" },
  primaryBtn: { border: "1px solid #1D4ED8", background: "#2563EB", color: "#FFFFFF", borderRadius: 12, padding: "10px 14px", fontSize: 13, fontWeight: 800, cursor: "pointer" },
  ghostBtn: { display: "inline-flex", alignItems: "center", gap: 6, border: "1px solid #D7DDEA", background: "#FFFFFF", color: "#1F2937", borderRadius: 12, padding: "9px 12px", fontSize: 13, fontWeight: 700, cursor: "pointer" },
  iconBtnDanger: { display: "inline-flex", alignItems: "center", justifyContent: "center", width: 36, height: 36, borderRadius: 10, border: "1px solid rgba(239, 68, 68, 0.28)", background: "rgba(254, 226, 226, 0.72)", color: "#DC2626", cursor: "pointer" },
  toggleRow: { display: "flex", alignItems: "flex-start", gap: 10, padding: "12px 14px", borderRadius: 14, border: "1px solid #E5E7EB", background: "#FAFBFF" },
  toggleTitle: { fontSize: 13, fontWeight: 800, color: "#111827", marginBottom: 2 },
  toolbarRow: { display: "flex", justifyContent: "flex-end" },
  listStack: { display: "grid", gap: 10 },
  inlineCard: { display: "flex", gap: 8, alignItems: "center" },
  blockCard: { display: "grid", gap: 10, padding: 12, borderRadius: 14, border: "1px solid #E5E7EB", background: "#FCFCFE" },
  emptyState: { borderRadius: 14, border: "1px dashed #D7DDEA", padding: "16px 14px", fontSize: 12, color: "#6B7280", textAlign: "center", background: "#FBFCFF" },
  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },
  localeRow: { display: "flex", flexWrap: "wrap", gap: 8 },
  localePill: { border: "1px solid #D7DDEA", background: "#FFFFFF", color: "#4B5563", borderRadius: 999, padding: "7px 11px", fontSize: 12, fontWeight: 700, cursor: "pointer" },
  localePillOn: { border: "1px solid #BFDBFE", background: "#DBEAFE", color: "#1D4ED8", borderRadius: 999, padding: "7px 11px", fontSize: 12, fontWeight: 800, cursor: "pointer" },
  signatureCard: { display: "grid", gap: 14, padding: 14, borderRadius: 16, border: "1px solid #E5E7EB", background: "#FCFCFE" },
  previewWrap: { display: "grid", gap: 10, padding: 12, borderRadius: 14, border: "1px dashed #D7DDEA", background: "#FFFFFF" },
  autoLabelRow: { display: "flex", gap: 10, alignItems: "flex-start", padding: "10px 12px", borderRadius: 12, border: "1px solid #E5E7EB", background: "#FCFCFE" },
  autoLabelBody: { display: "grid", gap: 4 },
  autoLabelChip: { display: "inline-flex", width: "fit-content", padding: "4px 10px", borderRadius: 8, fontSize: 12, fontWeight: 800, color: "#374151" },
  fontPreview: { borderRadius: 14, border: "1px solid #E5E7EB", background: "#FFFFFF", padding: 18 },
  loadingRoot: { minHeight: "100vh", display: "grid", placeItems: "center", background: "#F7F8FC", fontFamily: "\"Segoe UI\", system-ui, sans-serif" },
  loadingCard: { display: "inline-flex", alignItems: "center", gap: 10, padding: "12px 16px", borderRadius: 14, border: "1px solid #E5E7EB", background: "#FFFFFF", color: "#1F2937", fontSize: 13, fontWeight: 700 },
};
