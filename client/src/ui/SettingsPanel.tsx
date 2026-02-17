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

type Section = "general" | "ai" | "signature";

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
          applySkin((s as any).skinId || "classic");
        } catch {
          /* ignore */
        }
        try {
          applySkin((s as any).skinId || "classic");
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

  const title = useMemo(() => {
    if (section === "general") return "Geral";
    if (section === "ai") return "IA knowledge";
    return "Assinatura";
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
            Repor
          </button>
          <button style={S.btn} onClick={onSave} disabled={saving}>
            {saving ? "A guardar…" : "Guardar"}
          </button>
        </div>
      </div>

      <div style={S.card}>
        <div style={S.sidebar}>
          <button style={section === "general" ? S.sideItemOn : S.sideItem} onClick={() => setSection("general")}>
            Geral
          </button>
          <button style={section === "ai" ? S.sideItemOn : S.sideItem} onClick={() => setSection("ai")}>
            IA knowledge
          </button>
          <button style={section === "signature" ? S.sideItemOn : S.sideItem} onClick={() => setSection("signature")}>
            Assinatura
          </button>
        </div>

        <div style={S.content}>
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
                <div style={S.hint}>Classic mantém o visual atual. MailMaestro torna a UI mais compacta e limpa.</div>
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
            <div style={{ display: "grid", gap: 10 }}>
              <div style={S.hint}>Notas permanentes para a IA (ex.: regras da empresa, frases padrão, etc.).</div>
              <textarea
                style={S.textarea}
                value={(model.aiKnowledge || []).join("\n")}
                onChange={(e) =>
                  setModel({ ...model, aiKnowledge: e.target.value.split(/\r?\n/).map((s) => s.trim()).filter(Boolean) })
                }
                placeholder="Uma nota por linha…"
              />
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
                        placeholder='Ex.: <div>Com os melhores cumprimentos,<br/>Pedro Lopes<br/>DIVITEK</div>'
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
                        placeholder={"Ex.:\nCom os melhores cumprimentos,\nPedro Lopes\nDIVITEK"}
                      />
                    </div>
                  </div>
                );
              })}
            </div>
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
  hTitle: { fontWeight: 800, fontSize: 14 },

  card: {
    borderRadius: 16,
    background: "#ffffff",
    border: "1px solid #e6e8ef",
    boxShadow: "0 10px 30px rgba(20, 26, 52, 0.08)",
    padding: 10,
    display: "grid",
    gridTemplateColumns: "110px 1fr",
    gap: 10,
  },
  sidebar: {
    display: "grid",
    gap: 6,
    alignContent: "start",
  },
  content: {
    minHeight: 220,
  },

  sideItem: {
    borderRadius: 10,
    padding: "8px 10px",
    border: "1px solid transparent",
    background: "transparent",
    fontSize: 12,
    textAlign: "left",
    cursor: "pointer",
    color: "#1d2b4f",
  },
  sideItemOn: {
    borderRadius: 10,
    padding: "8px 10px",
    border: "1px solid #d7dbeb",
    background: "#f2f6ff",
    fontSize: 12,
    textAlign: "left",
    cursor: "pointer",
    color: "#0b2e7a",
    fontWeight: 700,
  },

  fieldLabel: {
    fontSize: 12,
    fontWeight: 700,
    color: "#2a3558",
    marginBottom: 6,
  },
  select: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid #d7dbeb",
    padding: "8px 10px",
    fontSize: 12,
    outline: "none",
  },
  input: {
    width: "100%",
    borderRadius: 10,
    border: "1px solid #d7dbeb",
    padding: "8px 10px",
    fontSize: 12,
    outline: "none",
  },
  textarea: {
    width: "100%",
    minHeight: 80,
    borderRadius: 12,
    border: "1px solid #d7dbeb",
    padding: 10,
    fontSize: 12,
    outline: "none",
    resize: "vertical",
  },
  hint: {
    fontSize: 11,
    color: "#66719a",
    lineHeight: 1.35,
  },

  btn: {
    borderRadius: 999,
    border: "1px solid #123a8f",
    background: "#123a8f",
    color: "#fff",
    padding: "6px 12px",
    fontSize: 12,
    fontWeight: 700,
    cursor: "pointer",
  },
  btnGhost: {
    borderRadius: 999,
    border: "1px solid #d7dbeb",
    background: "#ffffff",
    color: "#20315d",
    padding: "6px 12px",
    fontSize: 12,
    fontWeight: 700,
    cursor: "pointer",
  },

  okBox: {
    marginTop: 10,
    borderRadius: 12,
    padding: 10,
    fontSize: 12,
    border: "1px solid #cfe7d2",
    background: "#f1fbf2",
    color: "#255d2b",
  },
  errorBox: {
    marginTop: 10,
    borderRadius: 12,
    padding: 10,
    fontSize: 12,
    border: "1px solid #f3c0c0",
    background: "#fff2f2",
    color: "#8a1f1f",
  },
  note: { fontSize: 12, color: "#66719a" },
  error: { fontSize: 12, color: "#8a1f1f" },
};
