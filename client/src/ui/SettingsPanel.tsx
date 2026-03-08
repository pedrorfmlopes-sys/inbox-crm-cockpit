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
  type LangOption,
  type ReplyLength,
  type SkinId,
} from "../settings";
import { applySkin } from "./skins";
import * as Icons from "./icons";
import { useCockpit } from "../components/shell/CockpitProvider";
import { aiListModels } from "../api";

type Section = "general" | "conns" | "ai" | "persona" | "signature" | "protection";

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

function localeShort(loc: AppLocale): string {
  if (loc === "pt-PT") return "PT";
  if (loc === "es-ES") return "ES";
  if (loc === "en-GB") return "EN";
  if (loc === "it-IT") return "IT";
  if (loc === "de-DE") return "DE";
  return loc;
}

export function SettingsPanel(): JSX.Element {
  const [loading, setLoading] = useState(true);
  const [saving, setSaving] = useState(false);
  const [status, setStatus] = useState<string | null>(null);
  const [section, setSection] = useState<Section>("general");
  const [model, setModel] = useState<CockpitSettingsV1 | null>(null);

  // local-only uploaded signature images (dataURL), per locale
  const [sigImgLocal, setSigImgLocal] = useState<Partial<Record<AppLocale, string>>>({});
  const [availableModels, setAvailableModels] = useState<{ openai: string[]; gemini: string[] }>({ openai: [], gemini: [] });
  const [fetchingModels, setFetchingModels] = useState(false);

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
    if (section === "conns") return "Ligações (Odoo & IA)";
    if (section === "ai") return "IA Knowledge";
    if (section === "persona") return "Minha Persona";
    if (section === "signature") return "Assinatura";
    return "Proteção (O Moat)";
  }, [section]);

  async function onSave() {
    if (!model) return;
    setSaving(true);
    setStatus(null);
    try {
      await saveSettings(model);
      setStatus("Guardado.");
      setTimeout(() => setStatus(null), 1800);
    } catch (e: any) {
      setStatus(e?.message || "Falha ao guardar");
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

      setStatus("Reposto para os valores por defeito.");
      setTimeout(() => setStatus(null), 2200);
    } catch (e: any) {
      setStatus(e?.message || "Falha ao repor");
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

  function addPreset() {
    if (!model) return;
    const newPreset = { id: `p${Date.now()}`, name: "Novo Modelo", prompt: "" };
    setModel({ ...model, responsePresets: [...(model.responsePresets || []), newPreset] });
  }

  function removePreset(id: string) {
    if (!model) return;
    setModel({ ...model, responsePresets: (model.responsePresets || []).filter(p => p.id !== id) });
  }

  function updatePreset(id: string, field: "name" | "prompt", value: string) {
    if (!model) return;
    setModel({
      ...model,
      responsePresets: (model.responsePresets || []).map(p => p.id === id ? { ...p, [field]: value } : p)
    });
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
          <button style={section === "ai" ? S.sideItemOn : S.sideItem} onClick={() => setSection("ai")}>
            IA Knowledge
          </button>
          <button style={section === "persona" ? S.sideItemOn : S.sideItem} onClick={() => setSection("persona")}>
            Minha Persona
          </button>
          <button style={section === "signature" ? S.sideItemOn : S.sideItem} onClick={() => setSection("signature")}>
            Assinatura
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
                  <div style={S.fieldLabel}>Modelos de Resposta (Presets)</div>
                  <button style={{ ...S.btnGhost, padding: "4px 10px", height: "auto" }} onClick={addPreset}>
                    <Icons.Settings size={12} style={{ marginRight: 4 }} />
                    Adicionar
                  </button>
                </div>
                <div style={{ ...S.hint, marginBottom: 12 }}>Cria atalhos para respostas frequentes ou instruções específicas.</div>

                <div style={{ display: "grid", gap: 10 }}>
                  {(model.responsePresets || []).map((p) => (
                    <div key={p.id} style={{ padding: 10, border: "1px solid var(--iccc-card-border)", borderRadius: 12, background: "rgba(255,255,255,0.02)" }}>
                      <div style={{ display: "flex", gap: 8, marginBottom: 8 }}>
                        <input
                          style={{ ...S.input, fontWeight: 700 }}
                          value={p.name}
                          onChange={(e) => updatePreset(p.id, "name", e.target.value)}
                          placeholder="Nome do Modelo (ex: Pedido NIF)"
                          title="Nome do modelo"
                        />
                        <button
                          style={{ ...S.btnGhost, borderColor: "#fca5a5", color: "#ef4444" }}
                          onClick={() => removePreset(p.id)}
                          title="Remover modelo"
                        >
                          <Icons.Trash size={12} />
                        </button>
                      </div>
                      <textarea
                        style={{ ...S.textarea, minHeight: 60 }}
                        value={p.prompt}
                        onChange={(e) => updatePreset(p.id, "prompt", e.target.value)}
                        placeholder="Instruções para a IA (ex: Agradece e pede o NIF de faturação)..."
                        title="Instruções do modelo"
                      />
                    </div>
                  ))}
                  {(!model.responsePresets || model.responsePresets.length === 0) && (
                    <div style={{ ...S.hint, textAlign: "center", padding: 20, border: "1px dashed var(--iccc-card-border)", borderRadius: 12 }}>
                      Nenhum modelo criado. Adiciona um acima para acelerar as tuas respostas.
                    </div>
                  )}
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

          {section === ("protection" as any) && (
            <ProtectionSettings />
          )}

          {status && <div style={status.startsWith("Falha") ? S.errorBox : S.okBox}>{status}</div>}
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
  setStatus: (s: string | null) => void,
  availableModels: { openai: string[]; gemini: string[] },
  fetchingModels: boolean,
  refreshModels: () => Promise<void>
}) {
  const { granularStatus, granularStatusDetails, checkConnectivity, login } = useCockpit();
  const [isTesting, setIsTesting] = useState(false);

  const handleTest = async () => {
    setIsTesting(true);
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
