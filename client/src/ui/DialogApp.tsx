import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  createOdoo,
  linkEmailToRecord,
  odooPing,
  readOdoo,
  searchOdoo,
  searchOdooDomain,
  setApiSessionToken,
  writeOdoo,
  aiGenerate,
  getLeadTypeFieldMeta,
  type OdooFieldMeta,
} from "@/api";

import DebugPanel from "@/ui/DebugPanel";
import { prepareReferencedRecordName } from "@/referenceCodes";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import * as Icons from "./icons"; // Import icons symmetrically with CrmCockpit

/**
 * OdooMemoryCheck: Proactively searches for open tasks contextually.
 */
function OdooMemoryCheck({ partnerId, projectId, fromEmail }: { partnerId?: number | null, projectId?: number | null, fromEmail?: string }) {
  const [count, setCount] = useState(0);

  useEffect(() => {
    (async () => {
      let activePartnerId = partnerId;

      // Se não temos partnerId, tentamos encontrar por email
      if (!activePartnerId && !projectId && fromEmail) {
        try {
          const partners = await searchOdooDomain("res.partner", [["email", "=", fromEmail]], ["id"], 1);
          if (partners?.length) activePartnerId = partners[0].id;
        } catch (e) {
          console.error("[OdooMemory] Partner lookup failed", e);
        }
      }

      if (!activePartnerId && !projectId) return setCount(0);

      try {
        const domain: any[] = [["stage_id.is_closed", "=", false]];
        if (projectId) domain.push(["project_id", "=", projectId]);
        else if (activePartnerId) domain.push(["partner_id", "=", activePartnerId]);

        const rows = await searchOdooDomain("project.task", domain, ["id"], 100);
        setCount(rows?.length || 0);
      } catch (e) {
        console.error("[OdooMemory] Task search failed", e);
        setCount(0);
      }
    })();
  }, [partnerId, projectId, fromEmail]);

  if (count === 0) return null;

  return (
    <div style={{ marginBottom: 16 }}>
      <div style={S.yellowGlass}>
        <Icons.AlertCircle size={12} />
        {count} TAREFAS EM ABERTO PARA ESTE CONTEXTO
      </div>
    </div>
  );
}

/**
 * AiAssistant: Analyzes bodyText to provide summary and actions.
 */
function AiAssistant({ bodyText, onAddAction, manualOnly }: { bodyText: string, onAddAction: (title: string, type: string) => void, manualOnly: boolean }) {
  const [loading, setLoading] = useState(false);
  const [data, setData] = useState<{ summary: string[], actions: string[] } | null>(null);
  const [error, setError] = useState<string | null>(null);

  const analyze = async () => {
    console.log("[IA] Analyze triggered. bodyText length:", bodyText?.length);
    if (!bodyText || bodyText.trim().length === 0) {
      console.warn("[IA] Cannot analyze: bodyText is empty.");
      return;
    }
    setLoading(true);
    setData(null);
    setError(null);
    try {
      const trimmedBody = bodyText.slice(0, 4500);
      const res = await aiGenerate({
        action: "summarize_actions",
        mode: "fast",
        locale: "auto",
        tone: "neutro",
        email: {
          subject: "", // Contexto já disponível no banner
          from: "",
          to: [],
          cc: [],
          bodyText: trimmedBody
        }
      });

      if (res.ok && res.data) {
        setData(res.data);
      } else if (res.ok && res.text) {
        try {
          const jsonStr = res.text.substring(res.text.indexOf("{"), res.text.lastIndexOf("}") + 1);
          const parsed = JSON.parse(jsonStr);
          setData(parsed);
        } catch {
          const lines = res.text.split("\n").filter(l => l.trim().length > 10).slice(0, 3);
          setData({ summary: lines, actions: [] });
        }
      } else {
        setError("Falha a reanalisar. Verifica ligação/limites.");
      }
    } catch (e) {
      console.error("[IA] Analysis failed", e);
      setError("Erro na análise. Tenta novamente mais tarde.");
    } finally {
      setLoading(false);
    }
  };

  useEffect(() => {
    if (!manualOnly && bodyText && !data && !loading) {
      analyze();
    }
  }, [bodyText, data, loading, manualOnly]);

  useEffect(() => {
    setData(null);
    setError(null);
    setLoading(false);
  }, [bodyText]);

  if (!data || !data.summary.length) {
    if (loading) return (
      <div style={{ padding: 10, borderRadius: 12, border: "1px dashed #d6def2", background: "rgba(255,255,255,0.42)" }}>
        <div style={{ fontSize: 10, fontWeight: 800, color: "#2563eb", marginBottom: 4, display: "flex", alignItems: "center", gap: "6px" }}>
          <Icons.RefreshCw size={10} className="animate-spin" />
          ASSISTENTE IA • A ANALISAR...
        </div>
      </div>
    );

    return (
      <div style={{ padding: 10, borderRadius: 12, border: "1px solid #d6def2", background: "rgba(255,255,255,0.42)", backdropFilter: "blur(12px)" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
            <Icons.Sparkles size={12} color="#2563eb" />
            <div style={{ fontSize: 10, fontWeight: 800, color: "#2563eb", textTransform: "uppercase" }}>Assistente IA</div>
          </div>
          <button style={S.secondaryBtn} onClick={analyze} title={manualOnly ? "Analisar o conteudo" : "Reanalisar o conteudo"}>
            <Icons.RefreshCw size={10} />
            {data ? "REANALISAR" : "ANALISAR"}
          </button>
        </div>
        <div style={{ fontSize: 11, color: bodyText ? "#BF2600" : "#777" }}>
          {bodyText ? (manualOnly ? "Clique em Analisar para processar o email." : "A analise automatica esta ativa para este formulario.") : "O conteudo do email ainda nao foi carregado."}
        </div>
        {bodyText && manualOnly && <div style={{ fontSize: 10, color: "#5E6C84", marginTop: 6 }}>A analise agora corre apenas por clique manual.</div>}
        {error && <div style={{ fontSize: 10, color: "#BF2600", marginTop: 8 }}><b>ERRO:</b> {error}</div>}
      </div>
    );
  }

  return (
    <div style={{ padding: 10, borderRadius: 12, border: "1px solid #d6def2", background: "rgba(255,255,255,0.42)", backdropFilter: "blur(12px)" }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
          <Icons.Sparkles size={12} color="#2563eb" />
          <div style={{ fontSize: 10, fontWeight: 800, color: "#2563eb", textTransform: "uppercase" }}>Assistente IA</div>
        </div>
        <button style={S.secondaryBtn} onClick={analyze} title="Reanalisar o conteúdo" disabled={loading}>
          <Icons.RefreshCw size={10} className={loading ? "animate-spin" : ""} />
          {loading ? "..." : "REANALISAR"}
        </button>
      </div>

      {error && <div style={{ fontSize: 10, color: "#BF2600", marginBottom: 8 }}><b>ERRO:</b> {error}</div>}

      <ul style={{ margin: 0, paddingLeft: 16, fontSize: 11, color: "#172B4D", marginBottom: 12 }}>
        {data.summary.map((s: string, i: number) => (
          <li key={i} style={{ marginBottom: 4 }}>{s.replace(/^(Olá|Bom dia|Boa tarde|Boa noite)[^,.]*[.,\s]*/i, "").trim()}</li>
        ))}
      </ul>

      {data.actions.length > 0 && (
        <div style={{ display: "flex", flexWrap: "wrap", gap: 6 }}>
          {data.actions.map((act: string, i: number) => (
            <button
              key={i}
              style={S.compactPrimaryBtn}
              onClick={() => onAddAction(act, "project.task")}
              title="Criar tarefa"
            >
              <Icons.Plus size={10} />
              {act.length > 16 ? act.substring(0, 15) + ".." : act}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
type Mode = "new" | "add" | "edit";
type Entity = "project.task" | "helpdesk.ticket" | "project.project" | "crm.lead" | "res.partner";
type DialogUi = "classic" | "v2";

/**
 * VerticalActionCascade: Ultra-compact glossy pill menu (Dialog Version).
 */
const VerticalActionCascade: React.FC<{ current: string; onSelect: (type: string) => void; disabled?: boolean }> = ({ current, onSelect, disabled }) => {
  const [isOpen, setIsOpen] = useState(false);
  const [hoveredBtn, setHoveredBtn] = useState<string | null>(null);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setIsOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  const items = [
    { label: "Tarefa", type: "project.task", icon: "📝" },
    { label: "Ticket", type: "helpdesk.ticket", icon: "🎫" },
    { label: "Lead", type: "crm.lead", icon: "🎯" },
    { label: "Projeto", type: "project.project", icon: "🏗️" },
    { label: "Contato", type: "res.partner", icon: "👤" },
  ];

  const currentLabel = items.find(i => i.type === current)?.label || current;

  const primaryStyle: React.CSSProperties = {
    ...S.primaryBtn,
    transition: "all 0.18s ease",
    ...(hoveredBtn === "main" ? {
      background: "linear-gradient(180deg, rgba(110,180,255,0.98) 0%, rgba(20,120,230,0.95) 100%)",
      boxShadow: "0 6px 14px rgba(0,100,210,0.5), inset 0 1px 0 rgba(255,255,255,0.7), inset 0 -1px 0 rgba(0,0,0,0.15)",
      transform: "translateY(-1px)",
    } : {}),
  };

  const secondaryStyle = (key: string): React.CSSProperties => ({
    ...S.secondaryBtn,
    transition: "all 0.18s ease",
    ...(hoveredBtn === key ? {
      background: "linear-gradient(180deg, rgba(230,240,255,0.98) 0%, rgba(195,215,248,0.95) 100%)",
      boxShadow: "0 6px 14px rgba(0,80,200,0.15), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
      transform: "translateY(-1px)",
    } : {}),
  });

  return (
    <div ref={ref} style={{ display: "flex", flexDirection: "column", gap: "4px", width: "94px" }}>
      <button
        style={primaryStyle}
        onClick={() => !disabled && setIsOpen(!isOpen)}
        onMouseEnter={() => setHoveredBtn("main")}
        onMouseLeave={() => setHoveredBtn(null)}
        disabled={disabled}
      >
        {currentLabel.toUpperCase()}
      </button>

      {isOpen && (
        <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
          {items.map(item => (
            <button
              key={item.type}
              style={secondaryStyle(item.type)}
              onMouseEnter={() => setHoveredBtn(item.type)}
              onMouseLeave={() => setHoveredBtn(null)}
              onClick={() => {
                onSelect(item.type);
                setIsOpen(false);
              }}
            >
              <span style={{ fontSize: "12px" }}>{item.icon}</span>
              {item.label.toUpperCase()}
            </button>
          ))}
        </div>
      )}
    </div>
  );
};

/**
 * AttachmentPicker: Select attachments using compact glass pills.
 */
function AttachmentPicker({ attachments, selected, onToggle }: {
  attachments: any[],
  selected: string[],
  onToggle: (name: string) => void
}) {
  if (!attachments?.length) return null;

  const selectedCount = selected.length;

  return (
    <div style={{ marginTop: 16 }}>
      <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8, marginBottom: 6, flexWrap: "wrap" }}>
        <div>
          <label style={S.labBlock}>ANEXOS PARA O ODOO ({attachments.length})</label>
          <div style={{ fontSize: 11, color: "#5E6C84", marginTop: 2 }}>
            Seleciona os ficheiros que queres associar ao registo.
          </div>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap" }}>
          <button
            type="button"
            style={S.secondaryBtn}
            onClick={() => attachments.forEach((att) => { if (!selected.includes(att.name)) onToggle(att.name); })}
            title="Selecionar todos os anexos"
          >
            TODOS
          </button>
          <button
            type="button"
            style={S.secondaryBtn}
            onClick={() => selected.forEach((name) => onToggle(name))}
            title="Limpar selecao"
            disabled={!selectedCount}
          >
            LIMPAR
          </button>
        </div>
      </div>

      <div style={{ fontSize: 11, color: "#2563eb", fontWeight: 700, marginBottom: 8 }}>
        {selectedCount
          ? `${selectedCount} anexo(s) seguem para o Odoo.`
          : "Nenhum anexo selecionado. O registo sera criado sem ficheiros."}
      </div>

      <div style={S.attachmentCardGrid}>
        {attachments.map(att => {
          const isSelected = selected.includes(att.name);
          const dataUrl = attachmentToPreviewUrl(att);
          const textPreview = getAttachmentTextPreview(att);
          return (
            <button
              type="button"
              key={att.name}
              onClick={() => onToggle(att.name)}
              style={isSelected ? { ...S.attachmentPreviewCard, ...S.attachmentPreviewCardActive } : S.attachmentPreviewCard}
              title={att.name}
            >
              <div style={S.attachmentPreviewHeader}>
                <span style={S.attachmentPreviewCheck}>{isSelected ? "✓" : ""}</span>
                <span style={S.attachmentPreviewName}>{att.name}</span>
              </div>

              <div style={S.attachmentPreviewBody}>
                {isImageAttachment(att) && dataUrl ? (
                  <img src={dataUrl} alt={att.name} style={S.attachmentPreviewImage} />
                ) : isPdfAttachment(att) && dataUrl ? (
                  <iframe title={att.name} src={`${dataUrl}#toolbar=0&navpanes=0&scrollbar=0`} style={S.attachmentPreviewFrame} />
                ) : textPreview ? (
                  <div style={S.attachmentPreviewText}>{textPreview}</div>
                ) : (
                  <div style={S.attachmentPreviewFallback}>{attachmentKindLabel(att)}</div>
                )}
              </div>

              <div style={S.attachmentPreviewMeta}>
                <span>{normalizeMimeLabel(att.contentType || att.mimetype || "") || "ficheiro"}</span>
                <span>{formatBytes(att.size)}</span>
              </div>
            </button>
          );
        })}
      </div>
    </div>
  );
}

type OdooPublishMode = "html" | "text" | "description" | "custom";
type OdooPublishState = {
  postToChatter: boolean;
  chatterMode: OdooPublishMode;
  customText: string;
};

function htmlToReadableTextClient(html?: string) {
  const raw = String(html || "").trim();
  if (!raw) return "";
  try {
    const doc = new DOMParser().parseFromString(raw, "text/html");
    return String(doc.body?.textContent || "")
      .replace(/\r/g, "")
      .replace(/\n{3,}/g, "\n\n")
      .trim();
  } catch {
    return raw.replace(/<[^>]+>/g, " ").replace(/\s+/g, " ").trim();
  }
}

function shortenPreview(text: string, max = 420) {
  const clean = String(text || "").replace(/\s+/g, " ").trim();
  if (!clean) return "";
  if (clean.length <= max) return clean;
  return `${clean.slice(0, max - 1)}…`;
}

function escapeHtmlClient(value: string) {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function plainTextToHtmlClient(text?: string) {
  const normalized = String(text || "")
    .replace(/\r/g, "")
    .replace(/\u00a0/g, " ")
    .replace(/[\t ]+\n/g, "\n")
    .replace(/\n{3,}/g, "\n\n")
    .trim();

  if (!normalized) return "";

  return normalized
    .split(/\n{2,}/)
    .map((block) => `<p style="margin: 0 0 10px 0;">${escapeHtmlClient(block).replace(/\n/g, "<br/>")}</p>`)
    .join("\n");
}

function sanitizeEmailHtmlForOdooClient(html?: string) {
  let cleaned = String(html || "").trim();
  if (!cleaned) return "";

  const bodyMatch = cleaned.match(/<body[^>]*>([\s\S]*?)<\/body>/i);
  if (bodyMatch?.[1]) cleaned = bodyMatch[1];

  cleaned = cleaned.replace(/<!--[\s\S]*?-->/g, "");
  cleaned = cleaned.replace(/<(script|style|meta|link|title|head|xml|svg|canvas|noscript|iframe)[^>]*>[\s\S]*?<\/\1>/gi, "");
  cleaned = cleaned.replace(/<img\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/<picture\b[^>]*>[\s\S]*?<\/picture>/gi, "");
  cleaned = cleaned.replace(/<source\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/<\/?(o:p|v:[^>\s]+|w:[^>\s]+)\b[^>]*>/gi, "");
  cleaned = cleaned.replace(/<(div|table|section)[^>]*(?:class|id)=["'][^"']*(?:gmail_signature|signature|x_signature|moz-signature|apple-mail-signature)[^"']*["'][^>]*>[\s\S]*?<\/\1>/gi, "");
  cleaned = cleaned.replace(/\s+on[a-z]+\s*=\s*(['"]).*?\1/gi, "");
  cleaned = cleaned.replace(/\s+on[a-z]+\s*=\s*[^\s>]+/gi, "");
  cleaned = cleaned.replace(/\s(href|src)\s*=\s*(['"])\s*javascript:[\s\S]*?\2/gi, "");
  cleaned = cleaned.replace(/<(\/?)(html|body)\b[^>]*>/gi, "");
  cleaned = cleaned.trim();

  return cleaned;
}

function buildDescriptionEditorHtml(emailHtml?: string, emailText?: string, currentValue?: string) {
  const existing = String(currentValue || "").trim();
  if (existing) {
    const cleanedExisting = sanitizeEmailHtmlForOdooClient(existing);
    if (cleanedExisting) return cleanedExisting;
    const existingText = htmlToReadableTextClient(existing);
    return plainTextToHtmlClient(existingText || existing);
  }

  const cleanedEmailHtml = sanitizeEmailHtmlForOdooClient(emailHtml);
  if (cleanedEmailHtml) return cleanedEmailHtml;

  return plainTextToHtmlClient(emailText);
}

function normalizeDescriptionEditorHtml(raw?: string) {
  const cleaned = sanitizeEmailHtmlForOdooClient(raw);
  if (cleaned) return cleaned;
  const fallbackText = htmlToReadableTextClient(raw);
  return plainTextToHtmlClient(fallbackText);
}

function normalizeMimeLabel(value: string) {
  const raw = String(value || "").trim().toLowerCase();
  if (!raw) return "";
  if (raw === "application/x-pdf") return "pdf";
  if (raw.startsWith("image/")) return raw.replace("image/", "");
  if (raw.startsWith("application/")) return raw.replace("application/", "");
  if (raw.startsWith("text/")) return raw.replace("text/", "");
  return raw;
}

function normalizeAttachmentMime(att: any) {
  return String(att?.contentType || att?.mimetype || "").trim().toLowerCase();
}

function isImageAttachment(att: any) {
  return normalizeAttachmentMime(att).startsWith("image/");
}

function isPdfAttachment(att: any) {
  const mime = normalizeAttachmentMime(att);
  return mime === "application/pdf" || mime === "application/x-pdf" || /\.pdf$/i.test(String(att?.name || ""));
}

function attachmentToPreviewUrl(att: any) {
  const content = String(att?.content || "").trim();
  if (!content) return "";
  if (content.startsWith("data:")) return content;
  const mime = normalizeAttachmentMime(att) || "application/octet-stream";
  return `data:${mime};base64,${content}`;
}

function getAttachmentTextPreview(att: any) {
  const mime = normalizeAttachmentMime(att);
  if (!mime.startsWith("text/")) return "";
  const content = String(att?.content || "").trim();
  if (!content) return "";
  try {
    const normalized = content.replace(/^data:[^,]+,/, "");
    const decoded = atob(normalized);
    return decoded.replace(/\s+/g, " ").trim().slice(0, 120) || "";
  } catch {
    return "";
  }
}

function attachmentKindLabel(att: any) {
  const name = String(att?.name || "").trim();
  const ext = name.includes(".") ? name.split(".").pop() : "";
  return ext ? ext.toUpperCase() : "FICHEIRO";
}

function formatBytes(value: any) {
  const size = Number(value || 0);
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function RichHtmlEditor({
  value,
  onChange,
  placeholder,
}: {
  value: string;
  onChange: (next: string) => void;
  placeholder: string;
}) {
  const editorRef = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const el = editorRef.current;
    if (!el) return;
    const normalized = String(value || "");
    if (document.activeElement === el) return;
    if (el.innerHTML !== normalized) {
      el.innerHTML = normalized;
    }
  }, [value]);

  const run = (command: string, commandValue?: string) => {
    editorRef.current?.focus();
    document.execCommand(command, false, commandValue);
    const next = normalizeDescriptionEditorHtml(editorRef.current?.innerHTML || "");
    onChange(next);
  };

  const handleLink = () => {
    const url = window.prompt("URL do link");
    if (!url) return;
    run("createLink", url);
  };

  const handleReset = () => {
    onChange("");
    if (editorRef.current) editorRef.current.innerHTML = "";
  };

  const isEmpty = !htmlToReadableTextClient(value);

  return (
    <div style={S.editorCard}>
      <div style={S.editorToolbar}>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={() => run("bold")} title="Negrito">
          B
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={() => run("italic")} title="Italico">
          I
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={() => run("underline")} title="Sublinhado">
          U
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={() => run("insertUnorderedList")} title="Lista">
          •
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={() => run("insertOrderedList")} title="Lista numerada">
          1.
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={handleLink} title="Link">
          Link
        </button>
        <button type="button" style={S.editorToolBtn} onMouseDown={(e) => e.preventDefault()} onClick={handleReset} title="Limpar">
          Limpar
        </button>
      </div>
      <div style={S.editorSurfaceWrap}>
        {isEmpty ? <div style={S.editorPlaceholder}>{placeholder}</div> : null}
        <div
          ref={editorRef}
          style={S.editorSurface}
          contentEditable
          suppressContentEditableWarning
          onInput={() => onChange(normalizeDescriptionEditorHtml(editorRef.current?.innerHTML || ""))}
          onBlur={() => onChange(normalizeDescriptionEditorHtml(editorRef.current?.innerHTML || ""))}
        />
      </div>
    </div>
  );
}

function buildOdooPreviewHtml({
  subject,
  fromName,
  fromEmail,
  receivedAtIso,
  emailWebLink,
  publishState,
  emailHtml,
  emailText,
  description,
}: {
  subject?: string;
  fromName?: string;
  fromEmail?: string;
  receivedAtIso?: string;
  emailWebLink?: string;
  publishState: OdooPublishState;
  emailHtml?: string;
  emailText?: string;
  description?: string;
}) {
  if (!publishState.postToChatter) {
    return `<p style="margin:0;color:#5E6C84;">O email será apenas ligado ao registo, sem nova mensagem no chatter.</p>`;
  }

  const fallbackText = String(emailText || htmlToReadableTextClient(emailHtml) || "").trim();
  let innerHtml = "";
  switch (publishState.chatterMode) {
    case "html":
      innerHtml = sanitizeEmailHtmlForOdooClient(emailHtml) || plainTextToHtmlClient(fallbackText);
      break;
    case "text":
      innerHtml = plainTextToHtmlClient(fallbackText);
      break;
    case "custom":
      innerHtml = plainTextToHtmlClient(publishState.customText);
      break;
    case "description":
    default:
      innerHtml = plainTextToHtmlClient(description || fallbackText);
      break;
  }

  const safeFrom = `${String(fromName || "").trim()}${fromEmail ? ` <${fromEmail}>` : ""}`.trim() || "(desconhecido)";
  return [
    `<div style="font-family: Arial, sans-serif; line-height: 1.5; color: #172B4D;">`,
    `<div style="border-left: 3px solid #714B67; padding-left: 12px; margin-bottom: 16px; color: #5E6C84;">`,
    subject ? `<p style="margin: 0 0 4px 0;"><b>Assunto:</b> ${escapeHtmlClient(subject)}</p>` : "",
    `<p style="margin: 0 0 4px 0;"><b>De:</b> ${escapeHtmlClient(safeFrom)}</p>`,
    receivedAtIso ? `<p style="margin: 0 0 4px 0;"><b>Data:</b> ${escapeHtmlClient(receivedAtIso)}</p>` : "",
    emailWebLink ? `<p style="margin: 0;"><a href="${escapeHtmlClient(emailWebLink)}" target="_blank" rel="noreferrer">Ver no Outlook</a></p>` : "",
    `</div>`,
    `<div style="overflow-x:auto;">${innerHtml || `<p style="margin:0;color:#5E6C84;">Não há conteúdo preparado para o Odoo.</p>`}</div>`,
    `</div>`,
  ].filter(Boolean).join("");
}

function getDefaultPublishState(mode: Mode, emailHtml?: string, emailText?: string): OdooPublishState {
  return {
    postToChatter: mode !== "edit",
    chatterMode: emailHtml ? "html" : (emailText ? "text" : "description"),
    customText: "",
  };
}

function buildLinkPayloadForRecord(
  ctx: Ctx,
  model: string,
  recordId: number,
  recordName: string,
  publishState: OdooPublishState,
  emailText: string,
  description: string,
) {
  const fallbackText = String(emailText || htmlToReadableTextClient(ctx.bodyHtml) || "").trim();
  const safeDescriptionHtml = normalizeDescriptionEditorHtml(description);
  const safeDescriptionText = htmlToReadableTextClient(safeDescriptionHtml);
  let bodyHtml = "";
  let bodyText = "";

  switch (publishState.chatterMode) {
    case "html":
      bodyHtml = String(ctx.bodyHtml || "").trim();
      bodyText = fallbackText;
      break;
    case "text":
      bodyText = fallbackText;
      break;
    case "custom":
      bodyText = String(publishState.customText || "").trim();
      break;
    case "description":
    default:
      bodyHtml = safeDescriptionHtml || plainTextToHtmlClient(fallbackText);
      bodyText = safeDescriptionText || fallbackText;
      break;
  }

  return {
    conversationId: ctx.conversationId,
    model,
    recordId,
    recordName,
    internetMessageId: ctx.internetMessageId,
    itemId: ctx.itemId,
    subject: ctx.subject,
    fromEmail: ctx.fromEmail,
    fromName: ctx.fromName,
    receivedAtIso: ctx.receivedAtIso,
    emailWebLink: ctx.emailWebLink,
    bodyHtml,
    bodyText,
    postToChatter: publishState.postToChatter,
  };
}

async function uploadSelectedAttachmentsToRecord(model: string, recordId: number, emailAtts: any[], selectedAtts: string[]) {
  for (const fileName of selectedAtts) {
    const att = (emailAtts || []).find((item: any) => item.name === fileName);
    if (!att?.content) continue;
    try {
      await createOdoo("ir.attachment", {
        name: att.name,
        datas: att.content,
        datas_fname: att.name,
        mimetype: att.contentType,
        res_model: model,
        res_id: recordId,
        type: "binary",
      });
    } catch (error) {
      console.error("Erro ao enviar anexo", att.name, error);
    }
  }
}

function OdooContentEditor({
  mode,
  subject,
  fromName,
  fromEmail,
  receivedAtIso,
  emailWebLink,
  emailHtml,
  emailText,
  description,
  onDescriptionChange,
  publishState,
  onPublishChange,
  attachments,
  selectedAttachments,
  selectedAttachmentCount,
  totalAttachmentCount,
}: {
  mode: Mode;
  subject?: string;
  fromName?: string;
  fromEmail?: string;
  receivedAtIso?: string;
  emailWebLink?: string;
  emailHtml?: string;
  emailText: string;
  description: string;
  onDescriptionChange: (value: string) => void;
  publishState: OdooPublishState;
  onPublishChange: (next: OdooPublishState) => void;
  attachments: any[];
  selectedAttachments: string[];
  selectedAttachmentCount: number;
  totalAttachmentCount: number;
}) {
  const cleanEmailText = useMemo(
    () => String(emailText || htmlToReadableTextClient(emailHtml) || "").trim(),
    [emailHtml, emailText],
  );

  const previewText = useMemo(() => {
    if (!publishState.postToChatter) return "O email será apenas ligado ao registo, sem nova mensagem no chatter.";
    switch (publishState.chatterMode) {
      case "html":
        return shortenPreview(cleanEmailText || "Odoo vai receber o HTML original do email com limpeza de segurança.");
      case "text":
        return shortenPreview(cleanEmailText);
      case "custom":
        return shortenPreview(publishState.customText || "Escreve o texto que queres publicar no chatter.");
      case "description":
      default:
        return shortenPreview(description || cleanEmailText);
    }
  }, [cleanEmailText, description, publishState]);

  const odooPreviewHtml = useMemo(
    () => buildOdooPreviewHtml({
      subject,
      fromName,
      fromEmail,
      receivedAtIso,
      emailWebLink,
      publishState,
      emailHtml,
      emailText,
      description,
    }),
    [subject, fromName, fromEmail, receivedAtIso, emailWebLink, publishState, emailHtml, emailText, description],
  );

  const attachmentCards = useMemo(() => {
    const source = Array.isArray(attachments) ? attachments : [];
    const picked = source.filter((att: any) => selectedAttachments.includes(att.name));
    return (picked.length ? picked : source).slice(0, 6);
  }, [attachments, selectedAttachments]);

  return (
    <div style={S.odooEditorCard}>
      <div style={S.odooEditorHeader}>
        <div>
          <div style={S.odooEditorTitle}>CONTEUDO ODOO</div>
          <div style={S.odooEditorHint}>Revê aqui o resultado final antes de criar ou atualizar o registo no Odoo.</div>
        </div>
        <div style={{ display: "flex", gap: 6, flexWrap: "wrap", justifyContent: "flex-end" }}>
          <button
            type="button"
            style={S.secondaryBtn}
            onClick={() => onDescriptionChange(cleanEmailText)}
            title="Preencher descricao com o texto limpo do email"
          >
            EMAIL
          </button>
          <button
            type="button"
            style={S.secondaryBtn}
            onClick={() => onDescriptionChange("")}
            title="Limpar descricao"
          >
            LIMPAR
          </button>
        </div>
      </div>

      <div style={S.odooEditorGrid}>
        <label style={S.odooToggleRow}>
          <input
            type="checkbox"
            checked={publishState.postToChatter}
            onChange={(e) => onPublishChange({ ...publishState, postToChatter: e.target.checked })}
          />
          <span>{mode === "edit" ? "Adicionar mensagem no chatter" : "Publicar email no chatter"}</span>
        </label>

        <div style={S.odooEditorSelectWrap}>
          <span style={S.odooEditorMiniLab}>Origem</span>
          <select
            style={S.sel}
            value={publishState.chatterMode}
            disabled={!publishState.postToChatter}
            onChange={(e) => onPublishChange({ ...publishState, chatterMode: e.target.value as OdooPublishMode })}
          >
            <option value="html">HTML original</option>
            <option value="text">Texto limpo</option>
            <option value="description">Descricao</option>
            <option value="custom">Personalizado</option>
          </select>
        </div>
      </div>

      {publishState.postToChatter && publishState.chatterMode === "custom" ? (
        <textarea
          style={{ ...S.ta, minHeight: 100, marginTop: 10 }}
          value={publishState.customText}
          onChange={(e) => onPublishChange({ ...publishState, customText: e.target.value })}
          placeholder="Escreve a mensagem que queres enviar para o chatter do Odoo..."
        />
      ) : (
        <>
          <div style={S.odooPreviewBox}>
            <div style={{ fontSize: 10, fontWeight: 800, textTransform: "uppercase", color: "#2563eb", marginBottom: 4 }}>
              Resumo
            </div>
            {previewText || "Nao ha conteudo preparado para o chatter."}
          </div>
          <div style={{ ...S.odooEditorMiniLab, marginTop: 10 }}>PREVIEW FINAL NO ODOO</div>
          <div style={S.odooHtmlPreview} dangerouslySetInnerHTML={{ __html: odooPreviewHtml }} />
        </>
      )}

      <div style={S.odooAttachmentHint}>
        {selectedAttachmentCount > 0
          ? `${selectedAttachmentCount} anexo(s) selecionado(s) seguem para o registo no Odoo.`
          : totalAttachmentCount > 0
            ? "Podes selecionar anexos abaixo para os associar ao registo."
            : "Este email nao tem anexos disponiveis nesta fase."}
      </div>

      {attachmentCards.length ? (
        <div style={{ marginTop: 10 }}>
          <div style={S.odooEditorMiniLab}>PREVIEW DE ANEXOS</div>
          <div style={S.odooAttachmentPreviewGrid}>
            {attachmentCards.map((att: any) => {
              const dataUrl = attachmentToPreviewUrl(att);
              const textPreview = getAttachmentTextPreview(att);
              return (
                <div key={att.name} style={S.odooAttachmentPreviewCard}>
                  <div style={S.odooAttachmentPreviewName}>{att.name}</div>
                  {isImageAttachment(att) && dataUrl ? (
                    <img src={dataUrl} alt={att.name} style={S.odooAttachmentPreviewImage} />
                  ) : isPdfAttachment(att) && dataUrl ? (
                    <iframe title={att.name} src={`${dataUrl}#toolbar=0&navpanes=0&scrollbar=0`} style={S.odooAttachmentPreviewFrame} />
                  ) : textPreview ? (
                    <div style={S.odooAttachmentPreviewText}>{textPreview}</div>
                  ) : (
                    <div style={S.odooAttachmentPreviewFallback}>{attachmentKindLabel(att)}</div>
                  )}
                  <div style={S.odooAttachmentPreviewMeta}>
                    <span>{normalizeMimeLabel(att.contentType || att.mimetype || "") || "ficheiro"}</span>
                    <span>{formatBytes(att.size)}</span>
                  </div>
                </div>
              );
            })}
          </div>
        </div>
      ) : null}
    </div>
  );
}

type Recipient = { name: string; email: string };

function parseRecipientsParam(raw: string): Recipient[] {
  if (!raw) return [];
  return raw
    .split(";")
    .map((part) => {
      const [name, email] = part.split("|");
      return { name: String(name || "").trim(), email: String(email || "").trim() };
    })
    .filter((r) => r.email);
}

function qp() {
  return new URLSearchParams(window.location.search);
}

function getMode(): Mode {
  const m = (qp().get("mode") || "new").toLowerCase();
  return m === "add" || m === "edit" ? (m as Mode) : "new";
}

function getDialogUi(): DialogUi {
  return (qp().get("ui") || "").trim().toLowerCase() === "v2" ? "v2" : "classic";
}

type Ctx = {
  conversationId: string;
  internetMessageId: string;
  itemId?: string;
  subject: string;
  fromEmail: string;
  fromName: string;
  receivedAtIso: string;
  bodyHtml?: string;
  emailWebLink?: string;

  toR: Recipient[];
  ccR: Recipient[];
};

function getCtxFromQuery(): Ctx {
  const p = qp();
  return {
    conversationId: p.get("conversationId") || "",
    internetMessageId: p.get("internetMessageId") || "",
    itemId: p.get("itemId") || "",
    subject: p.get("subject") || "",
    fromEmail: p.get("fromEmail") || "",
    fromName: p.get("fromName") || "",
    receivedAtIso: p.get("receivedAtIso") || p.get("receivedDateTimeIso") || "",
    emailWebLink: p.get("emailWebLink") || "",
    toR: parseRecipientsParam(p.get("toR") || ""),
    ccR: parseRecipientsParam(p.get("ccR") || ""),
  };
}

function closeDialog() {
  // @ts-ignore global
  if (typeof Office !== "undefined" && Office?.context?.ui?.messageParent) {
    // @ts-ignore global
    Office.context.ui.messageParent("close");
    return;
  }
  window.close();
}

function shortId(s: string, head = 10, tail = 8) {
  if (!s) return "—";
  if (s.length <= head + tail + 3) return s;
  return `${s.slice(0, head)}...${s.slice(-tail)}`;
}

async function copyToClipboard(text: string) {
  try {
    await navigator.clipboard.writeText(text);
  } catch {
    // fallback
    const ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    document.execCommand("copy");
    document.body.removeChild(ta);
  }
}

async function withReferenceCode(model: string, values: Record<string, any>) {
  if (typeof values.name !== "string") return values;
  const prepared = await prepareReferencedRecordName(model, values.name);
  if (!prepared.referenceCode || prepared.title === values.name) return values;
  return { ...values, name: prepared.title };
}

type TypeaheadPickerProps = {
  label: string;
  placeholder: string;
  model: string;
  fields?: string[];
  limit?: number;
  pickedId: number | null;
  pickedName: string;
  onPick: (it: any) => void;
  extraDomain?: (q: string) => any[];
  compact?: boolean;
};

function TypeaheadPicker({
  label,
  placeholder,
  model,
  fields = ["id", "name", "display_name"],
  limit = 15,
  pickedId,
  pickedName,
  onPick,
  extraDomain,
  compact = false,
}: TypeaheadPickerProps) {
  const [q, setQ] = useState("");
  const [items, setItems] = useState<any[]>([]);
  const [open, setOpen] = useState(false);
  const [busy, setBusy] = useState(false);
  const debounceRef = useRef<number | null>(null);

  const effectiveText = pickedId ? pickedName : q;

  async function load(query: string) {
    setBusy(true);
    try {
      if (extraDomain) {
        const domain = extraDomain(query);
        const rows = await searchOdooDomain(model, domain, fields, limit);
        setItems(Array.isArray(rows) ? rows : []);
      } else {
        const rows = await searchOdoo(model, query, limit);
        setItems(Array.isArray(rows) ? rows : []);
      }
    } finally {
      setBusy(false);
    }
  }

  function scheduleLoad(query: string) {
    if (debounceRef.current) window.clearTimeout(debounceRef.current);
    debounceRef.current = window.setTimeout(() => load(query), 250);
  }

  useEffect(() => {
    if (!open) return;
    // quando abre, carrega logo (mesmo vazio) para mostrar 10–15
    scheduleLoad(pickedId ? "" : q);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [open]);

  useEffect(() => {
    if (!open) return;
    if (pickedId) return; // quando já está selecionado, não pesquisa
    scheduleLoad(q);
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [q]);

  return (
    <div style={{ marginTop: compact ? 0 : 10, position: "relative" }}>
      <label style={S.labBlock}>{label}</label>

      <div style={{ display: "flex", gap: 8, alignItems: "center" }}>
        <input
          style={{ ...S.input, flex: 1, minWidth: 0, height: "32px" }}
          value={effectiveText}
          onFocus={() => setOpen(true)}
          onBlur={() => setTimeout(() => setOpen(false), 150)}
          onChange={(e) => {
            const v = e.target.value;
            if (pickedId) {
              // se começar a escrever, limpa seleção
              setQ(v);
              onPick({ id: null, name: "" });
            } else {
              setQ(v);
            }
            setOpen(true);
          }}
          placeholder={placeholder}
        />

        {pickedId ? (
          <button
            className="jira-ghost-button" style={compact ? S.compactActionBtn : S.btn2}
            onClick={() => {
              onPick({ id: null, name: "" });
              setQ("");
              setOpen(true);
              load("");
            }}
            title="Limpar seleção"
          >
            Limpar
          </button>
        ) : (
          <button style={S.btn} onClick={() => load(q)} disabled={busy} title="Forçar pesquisa">
            {busy ? "…" : "Pesquisar"}
          </button>
        )}
      </div>

      {pickedId && !compact ? (
        <div style={{ marginTop: 6, fontSize: 12, color: "#666" }}>
          Selecionado: {pickedName} (#{pickedId})
        </div>
      ) : null}

      {open && (items?.length || busy) ? (
        <div style={S.pickList}>
          {busy && !items.length ? (
            <div style={{ padding: 10, color: "#777", fontSize: 12 }}>A procurar…</div>
          ) : null}

          {items.map((it: any) => (
            <button
              key={it.id}
              style={S.pickItem}
              onMouseDown={(e) => e.preventDefault()}
              onClick={() => {
                onPick(it);
                setOpen(false);
                setQ("");
              }}
            >
              <b>{it.display_name || it.name || `#${it.id}`}</b>
              <span style={{ color: "#777" }}>#{it.id}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

function CompactTypeCard({ value }: { value: string }) {
  return (
    <div style={S.metaCard}>
      <div style={S.metaCardLabel}>TIPO</div>
      <div style={S.metaCardValue}>{value}</div>
    </div>
  );
}

function CompactDualPickerCard({
  title,
  children,
}: {
  title: string;
  children: React.ReactNode;
}) {
  return (
    <div style={S.metaCardLarge}>
      <div style={S.metaCardLabel}>{title}</div>
      <div style={S.compactFieldStack}>{children}</div>
    </div>
  );
}

function DescriptionWorkspace({
  title,
  hint,
  value,
  onChange,
  placeholder,
  emailHtml,
  emailText,
}: {
  title: string;
  hint: string;
  value: string;
  onChange: (next: string) => void;
  placeholder: string;
  emailHtml?: string;
  emailText?: string;
}) {
  return (
    <div style={S.descriptionCard}>
      <div style={S.sectionHeaderRow}>
        <div>
          <div style={S.sectionTitle}>{title}</div>
          <div style={S.sectionHint}>{hint}</div>
        </div>
        <button
          type="button"
          style={S.compactActionBtn}
          onClick={() => onChange(buildDescriptionEditorHtml(emailHtml, emailText))}
          title="Usar o corpo do email"
        >
          Usar email
        </button>
      </div>
      <RichHtmlEditor value={value} onChange={onChange} placeholder={placeholder} />
    </div>
  );
}

function DialogShellV2({
  mode,
  entity,
  ctx,
  status,
  onSelectEntity,
  children,
}: {
  mode: Mode;
  entity: Entity;
  ctx: Ctx;
  status: string | null;
  onSelectEntity: (next: Entity) => void;
  children: React.ReactNode;
}) {
  const metaTone = status && /erro|falha|indispon/i.test(status) ? "#C25100" : "#0C66E4";

  return (
    <div style={{ ...S.page, gap: 0 }}>
      <div style={{
        position: "sticky",
        top: 0,
        zIndex: 10,
        background: "linear-gradient(180deg, rgba(244,248,255,0.98) 0%, rgba(244,248,255,0.94) 100%)",
        borderBottom: "1px solid #dbe4f3",
        padding: "12px 14px 10px",
        display: "grid",
        gap: 10,
      }}>
        <div style={{ display: "flex", justifyContent: "space-between", gap: 12, alignItems: "start" }}>
          <div style={{ minWidth: 0 }}>
            <div style={{ fontSize: 10, fontWeight: 800, letterSpacing: "0.06em", textTransform: "uppercase", color: "#5E6C84" }}>
              CRM 2
            </div>
            <div style={{ fontSize: 18, fontWeight: 800, color: "#172B4D", lineHeight: 1.15 }}>
              {mode === "new" ? "Novo registo" : mode === "edit" ? "Editar registo" : "Ligar existente"}
            </div>
            <div style={{ fontSize: 11, color: "#5E6C84", marginTop: 2, lineHeight: 1.4 }}>
              Editor paralelo para testar uma versao mais limpa do fluxo CRM.
            </div>
          </div>
          <button type="button" style={S.btn3} onClick={closeDialog}>Fechar</button>
        </div>

        <div style={{ display: "grid", gridTemplateColumns: "1.3fr 0.9fr", gap: 8 }}>
          <div style={{
            border: "1px solid #DFE1E6",
            borderRadius: 10,
            background: "#FFFFFF",
            padding: "10px 12px",
            display: "grid",
            gap: 4,
            minWidth: 0,
          }}>
            <div style={{ fontSize: 10, fontWeight: 800, color: "#5E6C84", textTransform: "uppercase", letterSpacing: "0.05em" }}>
              Assunto de origem
            </div>
            <div style={{ fontSize: 13, fontWeight: 700, color: "#172B4D", lineHeight: 1.35 }}>
              {ctx.subject || "Sem assunto"}
            </div>
            <div style={{ fontSize: 11, color: "#42526E" }}>
              {ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail}>` : (ctx.fromEmail || "Sem remetente")}
            </div>
          </div>

          <div style={{
            border: "1px solid #DFE1E6",
            borderRadius: 10,
            background: "#FFFFFF",
            padding: "10px 12px",
            display: "grid",
            gap: 8,
          }}>
            <div style={{ fontSize: 10, fontWeight: 800, color: "#5E6C84", textTransform: "uppercase", letterSpacing: "0.05em" }}>
              Tipo de registo
            </div>
            {mode === "edit" ? (
              <div style={{ fontSize: 13, fontWeight: 700, color: "#172B4D" }}>{entity}</div>
            ) : (
              <VerticalActionCascade current={entity} onSelect={(next) => onSelectEntity(next as Entity)} />
            )}
          </div>
        </div>

        {status ? (
          <div style={{
            borderRadius: 10,
            background: "rgba(12,102,228,0.07)",
            border: `1px solid ${metaTone === "#C25100" ? "rgba(194,81,0,0.28)" : "rgba(12,102,228,0.18)"}`,
            color: metaTone,
            padding: "8px 10px",
            fontSize: 11,
            lineHeight: 1.35,
          }}>
            {status}
          </div>
        ) : null}
      </div>

      <div style={{ ...S.scrollBody, paddingTop: 10 }}>
        <div style={{
          border: "1px solid #DFE1E6",
          borderRadius: 12,
          background: "#FFFFFF",
          padding: "10px 12px 14px",
        }}>
          {children}
        </div>
      </div>
    </div>
  );
}
function CompactOdooContentEditor(props: {
  mode: Mode;
  publishState: OdooPublishState;
  onPublishChange: (next: OdooPublishState) => void;
  selectedAttachmentCount: number;
  totalAttachmentCount: number;
}) {
  return (
    <div style={S.odooEditorCard}>
      <div style={S.odooEditorHeader}>
        <div>
          <div style={S.odooEditorTitle}>CONTEUDO ODOO</div>
          <div style={S.odooEditorHint}>Descricao = editor acima. Chatter = historico do email no Odoo.</div>
        </div>
      </div>

      <div style={S.odooEditorGrid}>
        <label style={S.odooToggleRow}>
          <input
            type="checkbox"
            checked={props.publishState.postToChatter}
            onChange={(e) => props.onPublishChange({ ...props.publishState, postToChatter: e.target.checked })}
          />
          <span>{props.mode === "edit" ? "Adicionar mensagem no chatter" : "Publicar email no chatter"}</span>
        </label>

        <div style={S.odooEditorSelectWrap}>
          <span style={S.odooEditorMiniLab}>Origem</span>
          <select
            style={S.sel}
            value={props.publishState.chatterMode}
            disabled={!props.publishState.postToChatter}
            onChange={(e) => props.onPublishChange({ ...props.publishState, chatterMode: e.target.value as OdooPublishMode })}
          >
            <option value="html">HTML original</option>
            <option value="text">Texto limpo</option>
            <option value="description">Descricao</option>
            <option value="custom">Personalizado</option>
          </select>
        </div>
      </div>

      {props.publishState.postToChatter && props.publishState.chatterMode === "custom" ? (
        <textarea
          style={{ ...S.ta, minHeight: 88, marginTop: 10 }}
          value={props.publishState.customText}
          onChange={(e) => props.onPublishChange({ ...props.publishState, customText: e.target.value })}
          placeholder="Mensagem curta para o chatter..."
        />
      ) : (
        <div style={S.odooMiniSummary}>
          {props.publishState.postToChatter
            ? `O email vai para o chatter em modo ${props.publishState.chatterMode === "html" ? "HTML original" : props.publishState.chatterMode === "text" ? "Texto limpo" : props.publishState.chatterMode === "description" ? "Descricao" : "Personalizado"}.`
            : "O email sera apenas ligado ao registo, sem nova mensagem no chatter."}
        </div>
      )}

      <div style={S.odooAttachmentHint}>
        {props.selectedAttachmentCount > 0
          ? `${props.selectedAttachmentCount} anexo(s) selecionado(s) seguem para o registo no Odoo.`
          : props.totalAttachmentCount > 0
            ? "Seleciona abaixo os anexos que queres associar ao registo."
            : "Este email nao tem anexos disponiveis nesta fase."}
      </div>
    </div>
  );
}

export default function DialogApp() {
  const isDevRuntime = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
  const dialogUi = getDialogUi();
  const [mode, setMode] = useState<Mode>(() => getMode());
  const [editId, setEditId] = useState<string | null>(() => qp().get("recordId") || null);
  const [ctx, setCtx] = useState<Ctx>(() => getCtxFromQuery());
  const [showThread, setShowThread] = useState(false);
  const [entity, setEntity] = useState<Entity>(() => {
    const m = qp().get("model") || "";
    return (m as Entity) || "project.task";
  });
  const [status, setStatus] = useState<string | null>(null);
  const [apiReady, setApiReady] = useState(false);
  const [aiManualOnly, setAiManualOnly] = useState(true);

  const [fullBody, setFullBody] = useState("");
  const [emailAtts, setEmailAtts] = useState<any[]>([]);

  useEffect(() => {
    if (!isDevRuntime) return;

    const runScenario = (scenario: any) => {
      console.log("[Simulation] Injecting scenario", scenario.id);
      setMode(scenario.context.mode);
      setEditId(scenario.context.editId || null);
      setEntity(scenario.context.entity);
      setFullBody(scenario.bodyText || "");
      setEmailAtts(scenario.attachments || []);
      setCtx(scenario.context);
    };

    (window as any).icccRunScenario = runScenario;

    const handler = (e: any) => runScenario(e.detail);
    window.addEventListener("iccc:run-scenario", handler);
    return () => {
      delete (window as any).icccRunScenario;
      window.removeEventListener("iccc:run-scenario", handler);
    };
  }, [isDevRuntime]);

  useEffect(() => {
    const b = localStorage.getItem("ic_bridge_body") || "";
    const h = localStorage.getItem("ic_bridge_html") || "";
    const a = localStorage.getItem("ic_bridge_atts");
    if (b) setFullBody(b);
    if (h) setCtx((prev) => ({ ...prev, bodyHtml: h }));
    if (a) {
      try { setEmailAtts(JSON.parse(a)); } catch { }
    }
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const st = await getSettings();
        setAiManualOnly(st.aiManualOnly !== false);
        if (st.odooSessionToken) {
          setApiSessionToken(st.odooSessionToken);
        }
        applySkin(st.skinId || 'classic');
      } catch {
        applySkin('classic');
      }
    })();
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const st = await getSettings();
        setAiManualOnly(st.aiManualOnly !== false);
        if (st.odooSessionToken) {
          setApiSessionToken(st.odooSessionToken);
          const pingResult = await odooPing();
          if (!pingResult.ok) {
            // Session expired? Try auto-login
            if (st.odooUrl && st.odooLogin && st.odooPassword) {
              const { login: apiLogin } = await import("@/api");
              const resp = await apiLogin({
                url: st.odooUrl,
                db: st.odooDb,
                login: st.odooLogin,
                password: st.odooPassword
              });
              if (resp.ok) {
                setApiSessionToken(resp.token);
              }
            }
          }
        }
      } catch (e: any) {
        setStatus(`API/Proxy falhou: ${e?.message || e}`);
      } finally {
        setApiReady(true);
      }

      if (ctx.subject || ctx.fromEmail || ctx.conversationId) return;

      // @ts-ignore global
      const OfficeAny = typeof Office !== "undefined" ? Office : null;
      const item = OfficeAny?.context?.mailbox?.item;
      if (!item) return;

      const subject = item.subject || "";
      const from = item.from;
      const fromEmail = from?.emailAddress || "";
      const fromName = from?.displayName || "";
      const conversationId = item.conversationId || "";
      const internetMessageId = item.internetMessageId || "";
      const normalize = (arr: any): Recipient[] =>
        Array.isArray(arr)
          ? arr
            .map((r: any) => ({ name: String(r?.displayName || "").trim(), email: String(r?.emailAddress || "").trim() }))
            .filter((r: any) => r.email)
          : [];

      setCtx((c) => ({
        ...c,
        subject,
        fromEmail,
        fromName,
        conversationId,
        internetMessageId,
        toR: c.toR?.length ? c.toR : normalize(item.to),
        ccR: c.ccR?.length ? c.ccR : normalize(item.cc),
      }));

      // Fallback: Se o bridge falhou ou está vazio, tenta ler o corpo agora
      if (item.body?.getAsync) {
        item.body.getAsync("text", (r: any) => {
          if (r?.status === OfficeAny?.AsyncResultStatus.Succeeded && r.value) {
            console.log("[Dialog] Body text fallback success");
            setFullBody(r.value);
          }
        });
        item.body.getAsync("html", (r: any) => {
          if (r?.status === OfficeAny?.AsyncResultStatus.Succeeded && r.value) {
            console.log("[Dialog] Body HTML success");
            setCtx(c => ({ ...c, bodyHtml: r.value }));
          }
        });
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  const statusAlert = status ? <div style={S.alert}>{status}</div> : null;

  const formContent = (
    <>
      {mode === "add" ? (
        <AddExistingPanel entity={entity} ctx={ctx} onStatus={setStatus} />
      ) : entity === "project.task" ? (
        <TaskForm
          mode={mode}
          ctx={ctx}
          editId={editId}
          fullBody={fullBody}
          emailAtts={emailAtts}
          onStatus={setStatus}
          fromEmail={ctx.fromEmail}
          apiReady={apiReady}
        />
      ) : entity === "project.project" ? (
        <ProjectForm
          mode={mode}
          ctx={ctx}
          editId={editId}
          fullBody={fullBody}
          emailAtts={emailAtts}
          onStatus={setStatus}
          fromEmail={ctx.fromEmail}
          apiReady={apiReady}
        />
      ) : entity === "crm.lead" ? (
        <LeadForm
          mode={mode}
          ctx={ctx}
          editId={editId}
          fullBody={fullBody}
          emailAtts={emailAtts}
          onStatus={setStatus}
          fromEmail={ctx.fromEmail}
          apiReady={apiReady}
        />
      ) : entity === "res.partner" ? (
        <ContactHubForm mode={mode} ctx={ctx} editId={editId} onStatus={setStatus} />
      ) : entity === "helpdesk.ticket" ? (
        <HelpdeskTicketForm
          mode={mode}
          ctx={ctx}
          editId={editId}
          fullBody={fullBody}
          emailAtts={emailAtts}
          onStatus={setStatus}
        />
      ) : (
        <div style={S.alert}>Este fluxo ainda usa o editor atual.</div>
      )}
    </>
  );

  if (dialogUi === "v2" && mode !== "add") {
    return (
      <DialogShellV2
        mode={mode}
        entity={entity}
        ctx={ctx}
        status={status}
        onSelectEntity={setEntity}
      >
        {formContent}
      </DialogShellV2>
    );
  }


  return (
    <div style={S.page}>
      {/* FIXED HEADER */}
      <div style={S.top}>
        <div style={S.titleBlock}>
          <div style={S.h1}>INBOX CRM</div>
          <div style={S.h2}>{mode === "new" ? "NOVO ITEM" : mode === "add" ? "LIGAR EXISTENTE" : "EDITAR"}</div>
        </div>
      </div>

      {/* SCROLLABLE BODY */}
      <div style={S.scrollBody}>
        <div style={S.banner}>
          <div style={S.bannerRow}>
            <b style={S.bannerLab}>DE</b>
            <span style={S.bannerVal}>{ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail}>` : (ctx.fromEmail || "—")}</span>
          </div>
          <div style={S.bannerRow}>
            <b style={S.bannerLab}>ASSUNTO</b>
            <span style={S.bannerVal}>{ctx.subject || "—"}</span>
          </div>
          <div style={S.bannerRow}>
            <b style={S.bannerLab}>PARA</b>
            <span style={S.bannerVal}>{ctx.toR?.length ? ctx.toR.map((r) => r.email).join("; ") : "—"}</span>
          </div>

          <div style={{ color: "#999", fontSize: 11, marginTop: 8, display: "flex", gap: 8, alignItems: "center" }}>
            {showThread ? (
              <>
                <span>Thread:</span>
                <code title={ctx.conversationId || ""} style={{ fontSize: 10 }}>{shortId(ctx.conversationId)}</code>
                {ctx.conversationId ? (
                  <button style={S.btn3} onClick={() => copyToClipboard(ctx.conversationId)}>Copiar</button>
                ) : null}
                <button style={S.threadToggle} onClick={() => setShowThread(false)} title="Ocultar thread">▴</button>
              </>
            ) : (
              <button style={S.threadToggle} onClick={() => setShowThread(true)} title="Mostrar thread">Thread ▾</button>
            )}
          </div>
        </div>

        <div style={S.formCard}>
          <div style={S.row}>
            <label style={S.lab}>TIPO</label>
            <VerticalActionCascade
              current={entity}
              onSelect={(type) => setEntity(type as Entity)}
              disabled={mode === "edit"}
            />
          </div>

          {formContent}
          {statusAlert}
        </div>
      </div>

      {/* FIXED FOOTER */}
      <div style={S.footer}>
        <div style={{ display: "flex", gap: 8 }}>
          <button className="jira-ghost-button" style={S.btn2} onClick={() => closeDialog()}>FECHAR</button>
        </div>
        <div style={{ color: "#6B778C", fontSize: 10, fontWeight: 700 }}>v6.2 • SPRINT 18 ULTRA-COMPACT GLOSSY</div>
      </div>

      {isDevRuntime ? <DebugPanel ctx={ctx} links={[]} meta={null} compact={true} /> : null}
    </div>
  );
}

function AddExistingPanel({ entity, ctx, onStatus }: any) {
  const [pickedId, setPickedId] = useState<number | null>(null);
  const [pickedName, setPickedName] = useState("");

  async function link() {
    if (!pickedId) return onStatus("Escolhe um registo para ligar.");
    try {
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: entity,
        recordId: pickedId,
        recordName: pickedName,
        internetMessageId: ctx.internetMessageId,
        itemId: ctx.itemId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
        bodyHtml: ctx.bodyHtml,
      });

      onStatus("Ligado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <TypeaheadPicker
        label="Selecionar existente"
        placeholder={`Pesquisar ${entity}...`}
        model={entity}
        pickedId={pickedId}
        pickedName={pickedName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPickedId(id);
          setPickedName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={link} disabled={!pickedId}>Ligar ao email</button>
        <button className="jira-ghost-button" style={S.btn2} onClick={() => closeDialog()}>Cancelar</button>
      </div>
    </div>
  );
}

function TaskForm({ mode, ctx, editId, onStatus, fullBody, emailAtts, fromEmail }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [description, setDescription] = useState("");
  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);
  const [publishState, setPublishState] = useState<OdooPublishState>(() => getDefaultPublishState(mode, ctx.bodyHtml, fullBody));
  const [projectId, setProjectId] = useState<number | null>(null);
  const [projectName, setProjectName] = useState("");
  const [assigneeId, setAssigneeId] = useState<number | null>(null);
  const [assigneeName, setAssigneeName] = useState("");
  const [deadline, setDeadline] = useState("");
  const [stageId, setStageId] = useState<number | null>(null);
  const [stageName, setStageName] = useState("");
  const [stagePick, setStagePick] = useState<any[]>([]);
  const [isSub, setIsSub] = useState(false);
  const [parentId, setParentId] = useState<number | null>(null);
  const [parentName, setParentName] = useState("");
  const [pendingSubtasks, setPendingSubtasks] = useState<string[]>([]);

  useEffect(() => {
    if (mode === "new" && !htmlToReadableTextClient(description) && (ctx.bodyHtml || fullBody)) {
      setDescription(buildDescriptionEditorHtml(ctx.bodyHtml, fullBody));
    }
  }, [mode, ctx.bodyHtml, fullBody, description]);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("project.task", [editId], ["name", "description", "project_id", "user_ids", "date_deadline", "stage_id", "parent_id"]);
        const record = rows?.[0];
        if (!record) return;
        setName(record.name || "");
        setDescription(record.description || "");
        if (record.project_id) { setProjectId(record.project_id[0]); setProjectName(record.project_id[1]); }
        if (Array.isArray(record.user_ids) && record.user_ids.length) {
          const users = await readOdoo("res.users", [record.user_ids[0]], ["name"]);
          setAssigneeId(record.user_ids[0]);
          setAssigneeName(users?.[0]?.name || "");
        }
        if (record.date_deadline) setDeadline(String(record.date_deadline));
        if (record.stage_id) { setStageId(record.stage_id[0]); setStageName(record.stage_id[1]); }
        if (record.parent_id) { setIsSub(true); setParentId(record.parent_id[0]); setParentName(record.parent_id[1]); }
      } catch (error: any) {
        onStatus(error?.message ?? String(error));
      }
    })();
  }, [editId, mode, onStatus]);

  useEffect(() => {
    (async () => {
      try {
        if (!projectId) return setStagePick([]);
        const rows = await searchOdooDomain("project.task.type", ["|", ["project_ids", "=", false], ["project_ids", "in", [projectId]]], ["id", "name"], 50);
        setStagePick(rows || []);
      } catch {
        setStagePick([]);
      }
    })();
  }, [projectId]);

  async function save() {
    try {
      let values: any = { name: name || "Nova tarefa", description: description || "" };
      if (projectId) values.project_id = projectId;
      if (assigneeId) values.user_ids = [[6, 0, [assigneeId]]];
      if (deadline) values.date_deadline = deadline;
      if (stageId) values.stage_id = stageId;
      if (isSub && parentId) values.parent_id = parentId;

      let id = editId;
      if (mode === "edit") {
        await writeOdoo("project.task", id, values);
        if (selectedAtts.length > 0) {
          onStatus("A enviar anexos...");
          await uploadSelectedAttachmentsToRecord("project.task", Number(id), emailAtts || [], selectedAtts);
        }
        if (publishState.postToChatter) {
          await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "project.task", Number(id), values.name, publishState, fullBody, description));
        }
        onStatus("Atualizado OK");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("project.task", values);
      id = await createOdoo("project.task", values);
      await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "project.task", id, values.name, publishState, fullBody, description));

      if (pendingSubtasks.length > 0) {
        onStatus(`A criar ${pendingSubtasks.length} subtarefas...`);
        for (const subTitle of pendingSubtasks) {
          try {
            const subtaskValues = await withReferenceCode("project.task", { name: subTitle, project_id: projectId || false, parent_id: id, user_ids: assigneeId ? [assigneeId] : [] });
            await createOdoo("project.task", subtaskValues);
          } catch (error) {
            console.error("Erro ao criar subtarefa diferida", error);
          }
        }
      }

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        await uploadSelectedAttachmentsToRecord("project.task", id, emailAtts || [], selectedAtts);
      }

      onStatus("Criado com sucesso");
      setTimeout(() => closeDialog(), 500);
    } catch (error: any) {
      onStatus(error?.message ?? String(error));
    }
  }

  async function handleAddAiAction(title: string) {
    if (mode === "new") {
      setPendingSubtasks((prev) => [...prev, title]);
      onStatus(`Subtarefa \"${title}\" agendada.`);
      return;
    }
    onStatus(`A criar tarefa: ${title}...`);
    try {
      let values: any = { name: title, project_id: projectId || false, user_ids: assigneeId ? [assigneeId] : [] };
      values = await withReferenceCode("project.task", values);
      const newId = await createOdoo("project.task", values);
      onStatus(`Tarefa criada (#${newId})`);
    } catch (error: any) {
      onStatus(`Falha IA: ${error.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck partnerId={projectId} projectId={projectId} fromEmail={fromEmail} />
      <div style={S.formTopGrid}>
        <div style={S.subjectCard}>
          <div style={S.metaCardLabel}>ASSUNTO</div>
          <input style={S.headerInput} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome da tarefa" />
        </div>
        <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} manualOnly={aiManualOnly} />
      </div>
      <div style={S.formMetaGrid}>
        <CompactTypeCard value="Tarefa" />
        <CompactDualPickerCard title="PROJETO E RESPONSAVEL">
          <TypeaheadPicker compact label="PROJETO" placeholder="Pesquisar projeto..." model="project.project" pickedId={projectId} pickedName={projectName} onPick={(item: any) => { const id = item?.id ?? null; setProjectId(id); setProjectName(id ? (item.display_name || item.name || `#${id}`) : ""); if (!id) { setStageId(null); setStageName(""); } }} />
          <TypeaheadPicker compact label="RESPONSAVEL" placeholder="Pesquisar utilizador..." model="res.users" fields={["id", "name", "display_name", "email"]} pickedId={assigneeId} pickedName={assigneeName} onPick={(item: any) => { const id = item?.id ?? null; setAssigneeId(id); setAssigneeName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        </CompactDualPickerCard>
      </div>
      <div style={S.grid2}>
        <PickerStatic label="ETAPA" pickedId={stageId} pickedName={stageName} items={stagePick} onPick={(item: any) => { setStageId(item.id); setStageName(item.name || item.display_name || `#${item.id}`); }} placeholder={projectId ? "Escolher etapa..." : "Etapa (opcional)"} />
        <div style={S.row}><label style={S.lab}>PRAZO</label><input style={S.input} type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} /></div>
      </div>
      <div style={S.row}><label style={S.lab}>SUBTAREFA</label><input type="checkbox" checked={isSub} onChange={(e) => { setIsSub(e.target.checked); if (!e.target.checked) { setParentId(null); setParentName(""); } }} /></div>
      {isSub ? <TypeaheadPicker label="PARENT TASK" placeholder={projectId ? "Pesquisar tarefa (filtra por projeto)..." : "Pesquisar tarefa (global)..."} model="project.task" fields={["id", "name", "display_name", "project_id"]} pickedId={parentId} pickedName={parentName} extraDomain={(query) => { const domain: any[] = []; if (projectId) domain.push(["project_id", "=", projectId]); if (query?.trim()) domain.push(["name", "ilike", query.trim()]); return domain; }} onPick={(item: any) => { const id = item?.id ?? null; setParentId(id); setParentName(id ? (item.display_name || item.name || `#${id}`) : ""); }} /> : null}
      <DescriptionWorkspace
        title="DESCRICAO"
        hint="Edita aqui a descricao estruturada da tarefa. Esta area alimenta a coluna esquerda do Odoo."
        value={description}
        onChange={setDescription}
        placeholder="Descricao e notas da tarefa..."
        emailHtml={ctx.bodyHtml}
        emailText={fullBody}
      />
      <CompactOdooContentEditor mode={mode} publishState={publishState} onPublishChange={setPublishState} selectedAttachmentCount={selectedAtts.length} totalAttachmentCount={(emailAtts || []).length} />
      <AttachmentPicker attachments={emailAtts} selected={selectedAtts} onToggle={(fileName) => setSelectedAtts((prev) => prev.includes(fileName) ? prev.filter((name) => name !== fileName) : [...prev, fileName])} />
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}><button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button></div>
    </div>
  );
}
function ProjectForm({ mode, ctx, editId, onStatus, fullBody, emailAtts, fromEmail }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [partnerId, setPartnerId] = useState<number | null>(null);
  const [partnerName, setPartnerName] = useState("");
  const [managerId, setManagerId] = useState<number | null>(null);
  const [managerName, setManagerName] = useState("");
  const [description, setDescription] = useState("");
  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);
  const [publishState, setPublishState] = useState<OdooPublishState>(() => getDefaultPublishState(mode, ctx.bodyHtml, fullBody));

  useEffect(() => {
    if (mode === "new" && !htmlToReadableTextClient(description) && (ctx.bodyHtml || fullBody)) {
      setDescription(buildDescriptionEditorHtml(ctx.bodyHtml, fullBody));
    }
  }, [mode, ctx.bodyHtml, fullBody, description]);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        let rows: any[] | null = null;
        try {
          rows = await readOdoo("project.project", [editId], ["name", "partner_id", "user_id", "description"]);
        } catch {
          rows = await readOdoo("project.project", [editId], ["name", "partner_id", "user_id"]);
        }
        const record = rows?.[0];
        if (!record) return;
        setName(record.name || "");
        if (record.partner_id) { setPartnerId(record.partner_id[0]); setPartnerName(record.partner_id[1]); }
        if (record.user_id) { setManagerId(record.user_id[0]); setManagerName(record.user_id[1]); }
        if (record.description) setDescription(String(record.description));
      } catch (error: any) {
        onStatus(error?.message ?? String(error));
      }
    })();
  }, [editId, mode, onStatus]);

  async function save() {
    try {
      let values: any = { name: name || "Novo projeto" };
      if (partnerId) values.partner_id = partnerId;
      if (managerId) values.user_id = managerId;
      if (description) values.description = description;

      if (mode === "edit") {
        try {
          await writeOdoo("project.project", editId, values);
        } catch {
          const fallbackValues = { ...values };
          delete fallbackValues.description;
          await writeOdoo("project.project", editId, fallbackValues);
        }
        if (selectedAtts.length > 0) {
          onStatus("A enviar anexos...");
          await uploadSelectedAttachmentsToRecord("project.project", Number(editId), emailAtts || [], selectedAtts);
        }
        if (publishState.postToChatter) {
          await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "project.project", Number(editId), values.name, publishState, fullBody, description));
        }
        onStatus("Atualizado OK");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("project.project", values);
      const id = await createOdoo("project.project", values);
      await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "project.project", id, values.name, publishState, fullBody, description));

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        await uploadSelectedAttachmentsToRecord("project.project", id, emailAtts || [], selectedAtts);
      }

      onStatus("Criado com sucesso");
      setTimeout(() => closeDialog(), 500);
    } catch (error: any) {
      onStatus(error?.message ?? String(error));
    }
  }

  async function handleAddAiAction(title: string) {
    onStatus(`A criar tarefa IA: ${title}...`);
    try {
      let values: any = { name: title, project_id: editId || false, partner_id: partnerId || false };
      values = await withReferenceCode("project.task", values);
      const newId = await createOdoo("project.task", values);
      onStatus(`Tarefa IA criada (#${newId})`);
    } catch (error: any) {
      onStatus(`Falha IA: ${error.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck partnerId={partnerId} fromEmail={fromEmail} />
      <div style={S.formTopGrid}>
        <div style={S.subjectCard}>
          <div style={S.metaCardLabel}>ASSUNTO</div>
          <input style={S.headerInput} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do projeto" />
        </div>
        <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} manualOnly={aiManualOnly} />
      </div>
      <div style={S.formMetaGrid}>
        <CompactTypeCard value="Projeto" />
        <CompactDualPickerCard title="CLIENTE E RESPONSAVEL">
          <TypeaheadPicker compact label="CLIENTE" placeholder="Pesquisar contacto/empresa..." model="res.partner" pickedId={partnerId} pickedName={partnerName} onPick={(item: any) => { const id = item?.id ?? null; setPartnerId(id); setPartnerName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
          <TypeaheadPicker compact label="RESPONSAVEL" placeholder="Pesquisar utilizador..." model="res.users" fields={["id", "name", "display_name", "email"]} pickedId={managerId} pickedName={managerName} onPick={(item: any) => { const id = item?.id ?? null; setManagerId(id); setManagerName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        </CompactDualPickerCard>
      </div>
      <DescriptionWorkspace
        title="DESCRICAO"
        hint="Prepara aqui a descricao limpa e formatada do projeto. Esta area alimenta a coluna esquerda do Odoo."
        value={description}
        onChange={setDescription}
        placeholder="Edita aqui a descricao do projeto..."
        emailHtml={ctx.bodyHtml}
        emailText={fullBody}
      />
      <CompactOdooContentEditor mode={mode} publishState={publishState} onPublishChange={setPublishState} selectedAttachmentCount={selectedAtts.length} totalAttachmentCount={(emailAtts || []).length} />
      <AttachmentPicker attachments={emailAtts} selected={selectedAtts} onToggle={(fileName) => setSelectedAtts((prev) => prev.includes(fileName) ? prev.filter((name) => name !== fileName) : [...prev, fileName])} />
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}><button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button></div>
    </div>
  );
}
function LeadForm({ mode, ctx, editId, onStatus, fullBody, emailAtts, fromEmail, apiReady }: any) {
  const LEAD_TYPE_FIELD_NAME = "x_studio_tipo_de_lead";
  const [name, setName] = useState(ctx.subject || "");
  const [contactName, setContactName] = useState(ctx.fromName || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");
  const [phone, setPhone] = useState("");
  const [partnerId, setPartnerId] = useState<number | null>(null);
  const [partnerName, setPartnerName] = useState("");
  const [stageId, setStageId] = useState<number | null>(null);
  const [stageName, setStageName] = useState("");
  const [description, setDescription] = useState("");
  const [leadTypeField, setLeadTypeField] = useState<OdooFieldMeta | null>(null);
  const [leadTypeLoading, setLeadTypeLoading] = useState(true);
  const [leadTypeError, setLeadTypeError] = useState<string | null>(null);
  const [leadTypeValue, setLeadTypeValue] = useState("");
  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);
  const [publishState, setPublishState] = useState<OdooPublishState>(() => getDefaultPublishState(mode, ctx.bodyHtml, fullBody));
  const leadTypeOptions = useMemo(() => Array.isArray(leadTypeField?.selection) ? leadTypeField.selection : [], [leadTypeField]);

  useEffect(() => {
    if (mode === "new" && !htmlToReadableTextClient(description) && (ctx.bodyHtml || fullBody)) {
      setDescription(buildDescriptionEditorHtml(ctx.bodyHtml, fullBody));
    }
  }, [mode, ctx.bodyHtml, fullBody, description]);

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        setLeadTypeLoading(true);
        setLeadTypeError(null);
        setLeadTypeField(null);
        const field = await getLeadTypeFieldMeta();
        if (!alive) return;
        setLeadTypeField(field);
      } catch (error: any) {
        if (alive) {
          setLeadTypeField(null);
          setLeadTypeError(error?.message || "Tipo de Lead indisponivel");
        }
      } finally {
        if (alive) setLeadTypeLoading(false);
      }
    })();
    return () => { alive = false; };
  }, [apiReady]);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        let rows: any[] | null = null;
        const baseFields = ["name", "contact_name", "email_from", "phone", "partner_id", "stage_id", LEAD_TYPE_FIELD_NAME];
        try {
          rows = await readOdoo("crm.lead", [editId], [...baseFields, "description"]);
        } catch {
          rows = await readOdoo("crm.lead", [editId], baseFields);
        }
        const record = rows?.[0];
        if (!record) return;
        setName(record.name || "");
        setContactName(record.contact_name || "");
        setEmail(record.email_from || "");
        setPhone(record.phone || "");
        if (record.partner_id) { setPartnerId(record.partner_id[0]); setPartnerName(record.partner_id[1]); }
        if (record.stage_id) { setStageId(record.stage_id[0]); setStageName(record.stage_id[1]); }
        if (record.description) setDescription(String(record.description));
        setLeadTypeValue(String(record[LEAD_TYPE_FIELD_NAME] || ""));
      } catch (error: any) {
        onStatus(error?.message ?? String(error));
      }
    })();
  }, [apiReady, editId, mode, onStatus]);

  async function save() {
    try {
      let values: any = {
        name: name || `Lead: ${ctx.subject || "sem assunto"}`,
        contact_name: contactName || "",
        email_from: email || "",
      };
      if (phone) values.phone = phone;
      if (partnerId) values.partner_id = partnerId;
      if (stageId) values.stage_id = stageId;
      if (description) values.description = description;
      values[LEAD_TYPE_FIELD_NAME] = leadTypeValue || false;

      if (mode === "edit") {
        try {
          await writeOdoo("crm.lead", editId, values);
        } catch {
          const fallbackValues = { ...values };
          delete fallbackValues.description;
          await writeOdoo("crm.lead", editId, fallbackValues);
        }
        if (selectedAtts.length > 0) {
          onStatus("A enviar anexos...");
          await uploadSelectedAttachmentsToRecord("crm.lead", Number(editId), emailAtts || [], selectedAtts);
        }
        if (publishState.postToChatter) {
          await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "crm.lead", Number(editId), values.name, publishState, fullBody, description));
        }
        onStatus("Atualizado OK");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("crm.lead", values);
      const id = await createOdoo("crm.lead", values);
      await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "crm.lead", id, values.name, publishState, fullBody, description));

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        await uploadSelectedAttachmentsToRecord("crm.lead", id, emailAtts || [], selectedAtts);
      }

      onStatus("Criado com sucesso");
      setTimeout(() => closeDialog(), 500);
    } catch (error: any) {
      onStatus(error?.message ?? String(error));
    }
  }

  async function handleAddAiAction(title: string) {
    onStatus(`A criar tarefa IA: ${title}...`);
    try {
      let values: any = { name: title, partner_id: partnerId || false };
      values = await withReferenceCode("project.task", values);
      const newId = await createOdoo("project.task", values);
      onStatus(`Tarefa IA criada (#${newId})`);
    } catch (error: any) {
      onStatus(`Falha IA: ${error.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck partnerId={partnerId} fromEmail={fromEmail} />
      <div style={S.formTopGrid}>
        <div style={S.subjectCard}>
          <div style={S.metaCardLabel}>ASSUNTO</div>
          <input style={S.headerInput} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do lead" />
        </div>
        <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} manualOnly={aiManualOnly} />
      </div>
      <div style={S.formMetaGrid}>
        <CompactTypeCard value="Lead" />
        <CompactDualPickerCard title="CLIENTE E RESPONSAVEL">
          <TypeaheadPicker compact label="EMPRESA" placeholder="Pesquisar res.partner..." model="res.partner" pickedId={partnerId} pickedName={partnerName} onPick={(item: any) => { const id = item?.id ?? null; setPartnerId(id); setPartnerName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
          <TypeaheadPicker compact label="ETAPA" placeholder="Pesquisar etapa do lead..." model="crm.stage" fields={["id", "name"]} pickedId={stageId} pickedName={stageName} onPick={(item: any) => { const id = item?.id ?? null; setStageId(id); setStageName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        </CompactDualPickerCard>
      </div>
      <div style={S.grid2}>
        <div style={S.row}><label style={S.lab}>CONTACTO</label><input style={S.input} value={contactName} onChange={(e) => setContactName(e.target.value)} placeholder="Nome do contacto" /></div>
        <div style={S.row}><label style={S.lab}>EMAIL</label><input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." /></div>
      </div>
      <div style={S.row}><label style={S.lab}>TELEFONE</label><input style={S.input} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefone" /></div>
      {leadTypeLoading && <div style={S.row}><label style={S.lab}>TIPO DE LEAD</label><select style={S.sel} value="" disabled><option value="">A carregar tipo de lead...</option></select></div>}
      {!leadTypeLoading && !!leadTypeField && <div style={S.row}><label style={S.lab}>{(leadTypeField.string || "Tipo de Lead").toUpperCase()}</label><select style={S.sel} value={leadTypeValue} onChange={(e) => setLeadTypeValue(e.target.value)}><option value="">Selecionar...</option>{leadTypeOptions.map(([value, label]) => <option key={value} value={value}>{label}</option>)}</select></div>}
      {!leadTypeLoading && !leadTypeField && <div style={S.row}><label style={S.lab}>TIPO DE LEAD</label><select style={S.sel} value="" disabled><option value="">{leadTypeError || "Tipo de Lead indisponivel"}</option></select></div>}
      <DescriptionWorkspace
        title="DESCRICAO"
        hint="Edita aqui o resumo estruturado do lead. Esta area alimenta a descricao do registo."
        value={description}
        onChange={setDescription}
        placeholder="Resumo e notas do lead..."
        emailHtml={ctx.bodyHtml}
        emailText={fullBody}
      />
      <CompactOdooContentEditor mode={mode} publishState={publishState} onPublishChange={setPublishState} selectedAttachmentCount={selectedAtts.length} totalAttachmentCount={(emailAtts || []).length} />
      <AttachmentPicker attachments={emailAtts} selected={selectedAtts} onToggle={(fileName) => setSelectedAtts((prev) => prev.includes(fileName) ? prev.filter((name) => name !== fileName) : [...prev, fileName])} />
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}><button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button></div>
    </div>
  );
}
function ContactHubForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.fromName || ctx.subject || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");
  const [phone, setPhone] = useState("");
  const [partnerKind, setPartnerKind] = useState<"person" | "company">("person");
  const [vat, setVat] = useState("");

  const participants = useMemo(() => {
    const out: Array<{ role: string; name: string; email: string }> = [];
    if (ctx.fromEmail) out.push({ role: "De", name: ctx.fromName || "", email: ctx.fromEmail });

    // Extract FROM, TO, CC
    (ctx.toR || []).forEach((r: any) => out.push({ role: "Para", name: r.name || "", email: r.email }));
    (ctx.ccR || []).forEach((r: any) => out.push({ role: "Cc", name: r.name || "", email: r.email }));

    // dedupe by email
    const seen = new Set<string>();
    return out.filter((p) => {
      if (!p.email) return false;
      const key = p.email.toLowerCase();
      if (seen.has(key)) return false;
      seen.add(key);
      return true;
    });
  }, [ctx]);

  const [match, setMatch] = useState<Record<string, { id: number; name: string; email?: string } | null>>({});

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("res.partner", [editId], ["name", "email", "phone", "company_type", "vat"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setEmail(r.email || "");
        setPhone(r.phone || "");
        setPartnerKind(r.company_type === "company" ? "company" : "person");
        setVat(String(r.vat || "").trim());
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  useEffect(() => {
    (async () => {
      const emails = participants.map((p) => p.email).filter(Boolean);
      const next: Record<string, any> = {};
      for (const em of emails) next[em] = null;
      setMatch(next);

      // lookup em série (simplicidade > performance nesta fase)
      for (const em of emails) {
        try {
          const rows = await searchOdooDomain("res.partner", [["email", "ilike", em]], ["id", "name", "display_name", "email"], 5);
          const found = rows?.find((r: any) => String(r.email || "").toLowerCase() === em.toLowerCase()) || rows?.[0];
          setMatch((prev) => ({ ...prev, [em]: found ? { id: found.id, name: found.display_name || found.name || `#${found.id}`, email: found.email } : null }));
        } catch {
          // ignore lookup errors
        }
      }
    })();
  }, [participants]);

  function normalizeVat(raw: string) {
    return String(raw || "").trim().toUpperCase().replace(/\s+/g, "");
  }

  async function findExistingCompanyByVat(rawVat: string) {
    const cleanVat = normalizeVat(rawVat);
    if (!cleanVat) return null;

    const domains: any[] = [
      [["company_type", "=", "company"], ["vat", "=", cleanVat]],
    ];
    if (/^\d{9}$/.test(cleanVat)) {
      domains.push([["company_type", "=", "company"], ["vat", "=", `PT${cleanVat}`]]);
    } else if (/^PT\d{9}$/.test(cleanVat)) {
      domains.push([["company_type", "=", "company"], ["vat", "=", cleanVat.slice(2)]]);
    }

    for (const domain of domains) {
      try {
        const rows = await searchOdooDomain("res.partner", domain, ["id", "name", "display_name", "vat"], 1);
        if (rows?.length) return rows[0];
      } catch {
        // ignore duplicate lookup errors
      }
    }

    return null;
  }

  async function saveMain() {
    try {
      const cleanName = String(name || "").trim();
      const cleanEmail = String(email || "").trim();
      const cleanPhone = String(phone || "").trim();
      const cleanVat = normalizeVat(vat);
      const isCompany = partnerKind === "company";

      if (isCompany && !cleanName) {
        onStatus("Indica o nome da empresa.");
        return;
      }
      if (isCompany && !cleanVat) {
        onStatus("Indica o NIF da empresa.");
        return;
      }

      if (mode === "edit") {
        const values: any = {
          name: cleanName || cleanEmail || (isCompany ? "Empresa" : "Contacto"),
          email: cleanEmail || false,
          phone: cleanPhone || false,
          company_type: partnerKind,
          is_company: isCompany,
          vat: isCompany ? cleanVat : false,
        };
        await writeOdoo("res.partner", editId, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      if (isCompany) {
        const existingCompany = await findExistingCompanyByVat(cleanVat);
        if (existingCompany?.id) {
          const display = existingCompany.display_name || existingCompany.name || `#${existingCompany.id}`;
          await linkEmailToRecord({
            conversationId: ctx.conversationId,
            model: "res.partner",
            recordId: existingCompany.id,
            recordName: display,
            internetMessageId: ctx.internetMessageId,
            itemId: ctx.itemId,
            subject: ctx.subject,
            fromEmail: ctx.fromEmail,
            fromName: ctx.fromName,
            receivedAtIso: ctx.receivedAtIso,
            emailWebLink: ctx.emailWebLink,
          });
          onStatus(`Empresa já existente ligada: ${display} ✅`);
          setTimeout(() => closeDialog(), 500);
          return;
        }
      }

      const values: any = {
        name: cleanName || cleanEmail || (isCompany ? `Empresa ${cleanVat}` : "Contacto"),
        company_type: partnerKind,
        is_company: isCompany,
      };
      if (cleanEmail) values.email = cleanEmail;
      if (cleanPhone) values.phone = cleanPhone;
      if (isCompany) values.vat = cleanVat;

      const id = await createOdoo("res.partner", values);
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "res.partner",
        recordId: id,
        recordName: cleanName || cleanEmail || (isCompany ? `Empresa ${cleanVat}` : `#${id}`),
        internetMessageId: ctx.internetMessageId,
        itemId: ctx.itemId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function linkToPartner(id: number, display: string) {
    try {
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "res.partner",
        recordId: id,
        recordName: display,
        internetMessageId: ctx.internetMessageId,
        itemId: ctx.itemId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });
      onStatus(`Ligado a ${display} ✅`);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function createPartnerFrom(p: any) {
    try {
      const id = await createOdoo("res.partner", { name: p.name || p.email, email: p.email });
      await linkToPartner(id, p.name || p.email);
      setMatch((prev) => ({ ...prev, [p.email]: { id, name: p.name || p.email, email: p.email } }));
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab}>TIPO</label>
        <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
          <button type="button" style={partnerKind === "person" ? S.btn : S.btn2} onClick={() => setPartnerKind("person")}>
            Pessoa
          </button>
          <button type="button" style={partnerKind === "company" ? S.btn : S.btn2} onClick={() => setPartnerKind("company")}>
            Empresa
          </button>
        </div>
      </div>

      <div style={S.row}>
        <label style={S.lab}>{partnerKind === "company" ? "EMPRESA" : "NOME"}</label>
        <input
          style={S.input}
          value={name}
          onChange={(e) => setName(e.target.value)}
          placeholder={partnerKind === "company" ? "Nome da empresa" : "Nome do contacto"}
        />
      </div>

      {partnerKind === "company" ? (
        <div style={S.row}>
          <label style={S.lab}>NIF</label>
          <input
            style={S.input}
            value={vat}
            onChange={(e) => setVat(e.target.value)}
            placeholder="NIF da empresa"
          />
        </div>
      ) : null}

      <div style={S.row}>
        <label style={S.lab}>EMAIL</label>
        <input
          style={S.input}
          value={email}
          onChange={(e) => setEmail(e.target.value)}
          placeholder={partnerKind === "company" ? "geral@empresa.pt (opcional)" : "email@..."}
        />
      </div>

      <div style={S.row}>
        <label style={S.lab}>TELEFONE</label>
        <input style={S.input} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefone" />
      </div>

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button style={S.btn} onClick={saveMain}>
          {mode === "edit" ? "Guardar" : "Criar"}
        </button>
      </div>

      <div style={{ marginTop: 20, borderTop: "1px solid #DFE1E6", paddingTop: 16 }}>
        <div style={{ fontWeight: 700, fontSize: 12, color: "#6B778C", marginBottom: 12, textTransform: "uppercase" }}>
          Participantes no email
        </div>

        {participants.length === 0 ? (
          <div style={{ color: "#5E6C84" }}>Sem participantes disponíveis.</div>
        ) : (
          <div style={{ display: "flex", flexDirection: "column", gap: 8 }}>
            {participants.map((p) => {
              const m = match[p.email];
              return (
                <div key={p.email} style={S.partRow}>
                  <div style={{ flex: 1, minWidth: 0 }}>
                    <div style={{ fontWeight: 600, fontSize: 13, color: "#172B4D", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" }}>
                      <span style={{ ...S.badge, background: "#DEEBFF", color: "#0747A6" }}>{p.role}</span> {p.name || p.email}
                    </div>
                    <div style={{ fontSize: 11, color: "#6B778C" }}>
                      {m ? `Odoo: ${m.name}` : "Não encontrado no Odoo"}
                    </div>
                  </div>

                  {m ? (
                    <button style={S.btn} onClick={() => linkToPartner(m.id, m.name)}>Ligar</button>
                  ) : (
                    <button style={S.btn} onClick={() => createPartnerFrom(p)}>Criar</button>
                  )}
                </div>
              );
            })}
          </div>
        )}
      </div>
    </div>
  );
}

function HelpdeskTicketForm({ mode, ctx, editId, onStatus, fullBody, emailAtts }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [description, setDescription] = useState("");
  const [partnerId, setPartnerId] = useState<number | null>(null);
  const [partnerName, setPartnerName] = useState("");
  const [teamId, setTeamId] = useState<number | null>(null);
  const [teamName, setTeamName] = useState("");
  const [assigneeId, setAssigneeId] = useState<number | null>(null);
  const [assigneeName, setAssigneeName] = useState("");
  const [stageId, setStageId] = useState<number | null>(null);
  const [stageName, setStageName] = useState("");
  const [priority, setPriority] = useState("0");
  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);
  const [publishState, setPublishState] = useState<OdooPublishState>(() => getDefaultPublishState(mode, ctx.bodyHtml, fullBody));

  useEffect(() => {
    if (mode === "new" && !htmlToReadableTextClient(description) && (ctx.bodyHtml || fullBody)) {
      setDescription(buildDescriptionEditorHtml(ctx.bodyHtml, fullBody));
    }
  }, [mode, ctx.bodyHtml, fullBody, description]);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("helpdesk.ticket", [editId], ["name", "description", "partner_id", "team_id", "user_id", "stage_id", "priority"]);
        const record = rows?.[0];
        if (!record) return;
        setName(record.name || "");
        setDescription(String(record.description || ""));
        if (record.partner_id) { setPartnerId(record.partner_id[0]); setPartnerName(record.partner_id[1]); }
        if (record.team_id) { setTeamId(record.team_id[0]); setTeamName(record.team_id[1]); }
        if (record.user_id) { setAssigneeId(record.user_id[0]); setAssigneeName(record.user_id[1]); }
        if (record.stage_id) { setStageId(record.stage_id[0]); setStageName(record.stage_id[1]); }
        setPriority(String(record.priority ?? "0"));
      } catch (error: any) {
        onStatus(error?.message ?? String(error));
      }
    })();
  }, [editId, mode, onStatus]);

  function handleAddAiAction(title: string) {
    setDescription((prev) => {
      const current = htmlToReadableTextClient(prev).trim();
      const nextText = current ? `${current}\n\n- ${title}` : `- ${title}`;
      return buildDescriptionEditorHtml(undefined, nextText);
    });
    onStatus(`Sugestao IA adicionada: ${title}`);
  }

  async function save() {
    try {
      let values: any = { name: name || `Ticket: ${ctx.subject || "sem assunto"}` };
      if (description) values.description = description;
      if (partnerId) values.partner_id = partnerId;
      if (teamId) values.team_id = teamId;
      if (assigneeId) values.user_id = assigneeId;
      if (stageId) values.stage_id = stageId;
      if (priority) values.priority = priority;

      let id = editId;
      if (mode === "edit") {
        await writeOdoo("helpdesk.ticket", id, values);
        if (selectedAtts.length > 0) {
          onStatus("A enviar anexos...");
          await uploadSelectedAttachmentsToRecord("helpdesk.ticket", Number(id), emailAtts || [], selectedAtts);
        }
        if (publishState.postToChatter) {
          await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "helpdesk.ticket", Number(id), values.name, publishState, fullBody, description));
        }
        onStatus("Atualizado OK");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("helpdesk.ticket", values);
      id = await createOdoo("helpdesk.ticket", values);
      await linkEmailToRecord(buildLinkPayloadForRecord(ctx, "helpdesk.ticket", id, values.name, publishState, fullBody, description));

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        await uploadSelectedAttachmentsToRecord("helpdesk.ticket", id, emailAtts || [], selectedAtts);
      }

      onStatus("Criado com sucesso");
      setTimeout(() => closeDialog(), 500);
    } catch (error: any) {
      onStatus(error?.message ?? String(error));
    }
  }

  return (
    <div>
      <div style={S.formTopGrid}>
        <div style={S.subjectCard}>
          <div style={S.metaCardLabel}>ASSUNTO</div>
          <input style={S.headerInput} value={name} onChange={(e) => setName(e.target.value)} placeholder="Titulo do ticket" />
        </div>
        <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} manualOnly={aiManualOnly} />
      </div>
      <div style={S.formMetaGrid}>
        <CompactTypeCard value="Ticket" />
        <CompactDualPickerCard title="CONTACTO E RESPONSAVEL">
          <TypeaheadPicker compact label="CONTACTO" placeholder="Pesquisar res.partner..." model="res.partner" pickedId={partnerId} pickedName={partnerName} onPick={(item: any) => { const id = item?.id ?? null; setPartnerId(id); setPartnerName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
          <TypeaheadPicker compact label="RESPONSAVEL" placeholder="Pesquisar utilizador..." model="res.users" fields={["id", "name", "display_name"]} pickedId={assigneeId} pickedName={assigneeName} onPick={(item: any) => { const id = item?.id ?? null; setAssigneeId(id); setAssigneeName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        </CompactDualPickerCard>
      </div>
      <div style={S.grid2}>
        <TypeaheadPicker label="EQUIPA" placeholder="Pesquisar equipa..." model="helpdesk.team" fields={["id", "name"]} pickedId={teamId} pickedName={teamName} onPick={(item: any) => { const id = item?.id ?? null; setTeamId(id); setTeamName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        <div />
      </div>
      <div style={S.grid2}>
        <TypeaheadPicker label="ETAPA" placeholder="Pesquisar etapa do ticket..." model="helpdesk.stage" fields={["id", "name"]} pickedId={stageId} pickedName={stageName} onPick={(item: any) => { const id = item?.id ?? null; setStageId(id); setStageName(id ? (item.display_name || item.name || `#${id}`) : ""); }} />
        <div style={S.row}><label style={S.lab}>PRIORIDADE</label><select style={S.sel} value={priority} onChange={(e) => setPriority(e.target.value)}><option value="0">Baixa</option><option value="1">Media</option><option value="2">Alta</option><option value="3">Urgente</option></select></div>
      </div>
      <DescriptionWorkspace
        title="DESCRICAO"
        hint="Edita aqui a descricao estruturada do ticket. O chatter fica separado e controlado abaixo."
        value={description}
        onChange={setDescription}
        placeholder="Detalhes e contexto do ticket..."
        emailHtml={ctx.bodyHtml}
        emailText={fullBody}
      />
      <CompactOdooContentEditor mode={mode} publishState={publishState} onPublishChange={setPublishState} selectedAttachmentCount={selectedAtts.length} totalAttachmentCount={(emailAtts || []).length} />
      <AttachmentPicker attachments={emailAtts} selected={selectedAtts} onToggle={(fileName) => setSelectedAtts((prev) => prev.includes(fileName) ? prev.filter((name) => name !== fileName) : [...prev, fileName])} />
      <div style={{ display: "flex", gap: 10, marginTop: 16 }}><button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button></div>
    </div>
  );
}
function GenericMiniForm({ mode, ctx, model, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.subject || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const fields =
          model === "res.partner" ? ["name", "email"] :
            model === "crm.lead" ? ["name", "email_from"] :
              ["name"];
        const rows = await readOdoo(model, [editId], fields);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        if (model === "res.partner") setEmail(r.email || "");
        if (model === "crm.lead") setEmail(r.email_from || "");
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  async function save() {
    try {
      if (mode === "edit") {
        const values: any = { name: name || "Atualizado" };
        if (model === "res.partner") values.email = email;
        if (model === "crm.lead") values.email_from = email;
        await writeOdoo(model, editId, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      let values: any =
        model === "res.partner" ? { name: name || email || "Novo contacto", email } :
          model === "crm.lead" ? { name: name || `Lead: ${ctx.subject || "sem assunto"}`, email_from: email } :
            model === "helpdesk.ticket" ? { name: name || `Ticket: ${ctx.subject || "sem assunto"}` } :
            { name: name || `Novo: ${ctx.subject || ""}` };

      values = await withReferenceCode(model, values);
      const id = await createOdoo(model, values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model,
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        itemId: ctx.itemId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab2}>{model === "helpdesk.ticket" ? "Título do ticket" : model === "crm.lead" ? "Nome do lead" : model === "res.partner" ? "Nome do contacto" : "Nome"}</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome" />
      </div>

      {(model === "crm.lead" || model === "res.partner") ? (
        <div style={S.row}>
          <label style={S.lab2}>Email</label>
          <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
        </div>
      ) : null}

      <div style={{ display: "flex", gap: 10, marginTop: 12 }}>
        <button style={S.btn} onClick={save}>{mode === "edit" ? "GUARDAR" : "CRIAR"}</button>
        <button className="jira-ghost-button" style={S.btn2} onClick={() => closeDialog()}>CANCELAR</button>
      </div>
    </div>
  );
}

function PickerStatic({ label, pickedId, pickedName, items, onPick, placeholder }: any) {
  const [open, setOpen] = useState(false);
  const ref = useRef<HTMLDivElement>(null);

  useEffect(() => {
    const handleClickOutside = (e: MouseEvent) => {
      if (ref.current && !ref.current.contains(e.target as Node)) setOpen(false);
    };
    document.addEventListener("mousedown", handleClickOutside);
    return () => document.removeEventListener("mousedown", handleClickOutside);
  }, []);

  return (
    <div ref={ref} style={{ marginTop: 10, position: "relative" }}>
      <label style={S.labBlock}>{label}</label>
      <div style={{ display: "flex", gap: 8 }}>
        <input
          style={{ ...S.input, cursor: "pointer" }}
          value={pickedId ? pickedName : ""}
          readOnly
          placeholder={placeholder}
          onClick={() => setOpen(!open)}
        />
      </div>
      {open && items?.length ? (
        <div style={{ ...S.pickList, maxHeight: "220px" }}>
          {items.map((it: any) => (
            <button key={it.id} style={S.pickItem} onClick={() => { onPick(it); setOpen(false); }}>
              <b>{it.display_name || it.name || `#${it.id}`}</b>
              <span style={{ color: "#777" }}>#{it.id}</span>
            </button>
          ))}
        </div>
      ) : null}
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  page: {
    fontFamily: "var(--iccc-font)",
    background: "linear-gradient(135deg, #f0f4f8 0%, #d9e8f5 50%, #e8edf5 100%)",
    height: "100vh",
    color: "#172B4D",
    display: "flex",
    flexDirection: "column",
    overflow: "hidden"
  },
  top: {
    padding: "12px 20px",
    borderBottom: "1px solid #DFE1E6",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    flexShrink: 0
  },
  h1: { fontWeight: 700, fontSize: 14, color: "#172B4D" },
  h2: { color: "#5E6C84", marginTop: 2, fontSize: 12 },

  scrollBody: {
    flex: 1,
    overflowY: "auto",
    padding: "16px 20px",
    background: "transparent",
  },

  banner: {
    background: "#F4F5F7",
    border: "1px solid #DFE1E6",
    borderRadius: 3,
    padding: "8px 12px",
    marginBottom: 16
  },
  bannerRow: {
    display: "grid",
    gridTemplateColumns: "70px 1fr",
    gap: 8,
    alignItems: "start",
    marginBottom: 4
  },
  bannerLab: { color: "#6B778C", fontSize: 11, fontWeight: 700, textTransform: "uppercase" },
  bannerVal: { fontSize: 13, color: "#172B4D", minWidth: 0, wordBreak: "break-all" },

  formCard: { background: "rgba(255,255,255,0.55)", backdropFilter: "blur(8px)", WebkitBackdropFilter: "blur(8px)", border: "1px solid rgba(255,255,255,0.4)", borderRadius: "12px", padding: "16px", marginBottom: 12 },
  formTopGrid: { display: "grid", gridTemplateColumns: "minmax(0, 1.25fr) minmax(280px, 0.75fr)", gap: 12, alignItems: "stretch", marginBottom: 12 },
  formMetaGrid: { display: "grid", gridTemplateColumns: "180px minmax(0, 1fr)", gap: 12, alignItems: "stretch", marginBottom: 12 },
  subjectCard: { border: "1px solid #d6def2", borderRadius: 12, background: "rgba(255,255,255,0.55)", padding: 10, display: "flex", flexDirection: "column", gap: 6, minHeight: 78 },
  metaCard: { border: "1px solid #d6def2", borderRadius: 12, background: "rgba(255,255,255,0.55)", padding: 10, display: "flex", flexDirection: "column", gap: 8, minHeight: 78 },
  metaCardLarge: { border: "1px solid #d6def2", borderRadius: 12, background: "rgba(255,255,255,0.55)", padding: 10, display: "flex", flexDirection: "column", gap: 8, minHeight: 78 },
  metaCardLabel: { fontSize: 10, fontWeight: 700, color: "#6B778C", textTransform: "uppercase", letterSpacing: "0.04em" },
  metaCardValue: { fontSize: 18, fontWeight: 700, color: "#172B4D" },
  compactFieldStack: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 10, alignItems: "start" },
  headerInput: { width: "100%", padding: "8px 10px", border: "2px solid #DFE1E6", borderRadius: 8, color: "#172B4D", background: "#FAFBFC", fontSize: 16, fontWeight: 600, outline: "none" },
  descriptionCard: { marginTop: 12, border: "1px solid #d6def2", borderRadius: 12, background: "rgba(255,255,255,0.55)", padding: 12 },
  sectionHeaderRow: { display: "flex", justifyContent: "space-between", gap: 10, alignItems: "flex-start", flexWrap: "wrap", marginBottom: 8 },
  sectionTitle: { fontSize: 11, fontWeight: 800, color: "#2563eb", textTransform: "uppercase" },
  sectionHint: { fontSize: 11, color: "#5E6C84", marginTop: 2 },
  editorCard: { border: "1px solid #d6def2", borderRadius: 10, background: "#fff", overflow: "hidden" },
  editorToolbar: { display: "flex", gap: 6, flexWrap: "wrap", padding: "8px", borderBottom: "1px solid #E6ECF7", background: "#F8FAFF" },
  editorToolBtn: { border: "1px solid #d6def2", borderRadius: 8, background: "#fff", color: "#253858", minWidth: 34, height: 28, padding: "0 8px", fontSize: 11, fontWeight: 600, cursor: "pointer" },
  editorSurfaceWrap: { position: "relative", minHeight: 260 },
  editorSurface: { minHeight: 260, padding: 12, color: "#172B4D", background: "#fff", fontSize: 13, lineHeight: 1.55, outline: "none", overflow: "auto" },
  editorPlaceholder: { position: "absolute", left: 12, top: 12, color: "#7A869A", fontSize: 13, pointerEvents: "none" },

  footer: {
    padding: "12px 20px",
    borderTop: "1px solid rgba(255,255,255,0.4)",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    background: "rgba(255,255,255,0.6)",
    backdropFilter: "blur(8px)",
    WebkitBackdropFilter: "blur(8px)",
    flexShrink: 0
  },

  row: { display: "grid", gridTemplateColumns: "100px 1fr", gap: 12, alignItems: "start", marginTop: 12 },
  lab: { fontSize: "12px", fontWeight: 700, color: "#6B778C", textTransform: "uppercase" },
  lab2: { fontSize: "12px", fontWeight: 700, color: "#6B778C", textTransform: "uppercase" },
  labBlock: { display: "block", fontSize: "12px", fontWeight: 700, marginBottom: 6, color: "#6B778C", textTransform: "uppercase" },

  sel: { padding: "6px 8px", border: "1px solid #DFE1E6", borderRadius: 3, color: "#172B4D", background: "#FAFBFC", fontSize: 13, height: "32px", outline: "none" },
  input: { width: "100%", padding: "6px 8px", border: "2px solid #DFE1E6", borderRadius: 3, color: "#172B4D", background: "#FAFBFC", fontSize: 13, height: "32px", outline: "none" },
  ta: { width: "100%", minHeight: 80, padding: "8px", border: "2px solid #DFE1E6", borderRadius: 8, resize: "vertical", color: "#172B4D", background: "#FAFBFC", fontSize: 13, outline: "none" },
  odooEditorCard: { marginTop: 12, padding: "12px", borderRadius: 12, border: "1px solid #d6def2", background: "rgba(255,255,255,0.55)" },
  odooEditorHeader: { display: "flex", justifyContent: "space-between", gap: 8, alignItems: "flex-start", flexWrap: "wrap" },
  odooEditorTitle: { fontSize: 11, fontWeight: 800, color: "#2563eb", textTransform: "uppercase" },
  odooEditorHint: { fontSize: 11, color: "#5E6C84", marginTop: 2 },
  odooEditorGrid: { display: "grid", gridTemplateColumns: "1fr 180px", gap: 10, marginTop: 10, alignItems: "end" },
  odooToggleRow: { display: "flex", alignItems: "center", gap: 8, fontSize: 12, color: "#172B4D", fontWeight: 600 },
  odooEditorSelectWrap: { display: "flex", flexDirection: "column", gap: 4 },
  odooEditorMiniLab: { fontSize: 10, fontWeight: 700, color: "#6B778C", textTransform: "uppercase" },
  odooMiniSummary: { marginTop: 10, padding: "9px 11px", borderRadius: 10, border: "1px solid #d6def2", background: "#F7F9FC", fontSize: 11, color: "#253858", lineHeight: 1.45 },
  odooPreviewBox: { marginTop: 10, padding: "10px 12px", borderRadius: 10, border: "1px solid #d6def2", background: "#F7F9FC", fontSize: 12, color: "#253858", lineHeight: 1.45, whiteSpace: "pre-wrap" },
  odooHtmlPreview: { marginTop: 10, padding: "12px", borderRadius: 10, border: "1px solid #d6def2", background: "#FFFFFF", color: "#172B4D", fontSize: 12, lineHeight: 1.5, maxHeight: 260, overflow: "auto" },
  odooAttachmentHint: { marginTop: 8, fontSize: 11, color: "#5E6C84" },
  odooAttachmentPreviewGrid: { marginTop: 8, display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 8 },
  odooAttachmentPreviewCard: { border: "1px solid #d6def2", borderRadius: 10, background: "#F7F9FC", padding: "8px", display: "flex", flexDirection: "column", gap: 6, minHeight: 138 },
  odooAttachmentPreviewName: { fontSize: 11, fontWeight: 700, color: "#172B4D", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  odooAttachmentPreviewImage: { width: "100%", height: 88, objectFit: "contain", borderRadius: 8, background: "#fff" },
  odooAttachmentPreviewFrame: { width: "100%", height: 88, border: "none", borderRadius: 8, background: "#fff" },
  odooAttachmentPreviewText: { fontSize: 11, color: "#253858", lineHeight: 1.35, background: "#fff", borderRadius: 8, padding: 8, minHeight: 88, overflow: "hidden" },
  odooAttachmentPreviewFallback: { display: "flex", alignItems: "center", justifyContent: "center", minHeight: 88, borderRadius: 8, background: "#fff", fontSize: 12, fontWeight: 700, color: "#5E6C84" },
  odooAttachmentPreviewMeta: { display: "flex", justifyContent: "space-between", gap: 8, fontSize: 10, color: "#6B778C" },
  attachmentCardGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(126px, 1fr))", gap: "8px", marginTop: "4px" },
  attachmentPreviewCard: { border: "1px solid #d6def2", borderRadius: 10, background: "#F7F9FC", padding: "8px", display: "flex", flexDirection: "column", gap: 6, minHeight: 118, cursor: "pointer", textAlign: "left" as const },
  attachmentPreviewCardActive: { border: "1px solid #2563eb", boxShadow: "0 0 0 2px rgba(37,99,235,0.12) inset", background: "#EEF4FF" },
  attachmentPreviewHeader: { display: "flex", alignItems: "center", gap: 6 },
  attachmentPreviewCheck: { width: 16, height: 16, borderRadius: 999, border: "1px solid #c3d4f4", display: "inline-flex", alignItems: "center", justifyContent: "center", fontSize: 11, fontWeight: 800, color: "#2563eb", flexShrink: 0 },
  attachmentPreviewName: { fontSize: 10, fontWeight: 700, color: "#172B4D", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  attachmentPreviewBody: { borderRadius: 8, background: "#fff", minHeight: 64, overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center" },
  attachmentPreviewImage: { width: "100%", height: 64, objectFit: "contain" },
  attachmentPreviewFrame: { width: "100%", height: 64, border: "none" },
  attachmentPreviewText: { fontSize: 10, color: "#253858", lineHeight: 1.3, padding: 6 },
  attachmentPreviewFallback: { fontSize: 12, fontWeight: 700, color: "#5E6C84" },
  attachmentPreviewMeta: { display: "flex", justifyContent: "space-between", gap: 8, fontSize: 10, color: "#6B778C" },

  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },

  btn: {
    boxSizing: "border-box",
    width: "auto", minWidth: "78px", maxWidth: "140px",
    height: "24px", minHeight: "24px", maxHeight: "24px",
    borderRadius: "16px",
    border: "1px solid rgba(0, 80, 180, 0.4)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    textTransform: "none",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(80,160,255,0.95) 0%, rgba(0,100,210,0.85) 100%)",
    color: "#FFFFFF",
    boxShadow: "0 4px 10px rgba(0,100,210,0.35), inset 0 1px 0 rgba(255,255,255,0.55), inset 0 -1px 0 rgba(0,0,0,0.15)",
  },
  btn2: {
    boxSizing: "border-box",
    width: "auto", minWidth: "78px", maxWidth: "140px",
    height: "24px", minHeight: "24px", maxHeight: "24px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    textTransform: "none",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
    color: "#172B4D",
    boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
  },
  btn3: {
    boxSizing: "border-box",
    width: "auto", minWidth: "78px", maxWidth: "140px",
    height: "24px", minHeight: "24px", maxHeight: "24px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    textTransform: "none",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
    color: "#172B4D",
    boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
  },

  alert: { marginTop: 12, padding: "8px 12px", borderRadius: 3, border: "1px solid #FFBDAD", background: "#FFEBE6", color: "#BF2600", fontSize: 12 },

  primaryBtn: {
    boxSizing: "border-box",
    width: "auto", minWidth: "78px", maxWidth: "140px",
    height: "24px", minHeight: "24px", maxHeight: "24px",
    borderRadius: "16px",
    border: "1px solid rgba(0, 80, 180, 0.4)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    textTransform: "none",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(80,160,255,0.95) 0%, rgba(0,100,210,0.85) 100%)",
    color: "#FFFFFF",
    boxShadow: "0 4px 10px rgba(0,100,210,0.35), inset 0 1px 0 rgba(255,255,255,0.55), inset 0 -1px 0 rgba(0,0,0,0.15)",
  },
  secondaryBtn: {
    boxSizing: "border-box",
    width: "auto", minWidth: "78px", maxWidth: "140px",
    height: "24px", minHeight: "24px", maxHeight: "24px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    textTransform: "none",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
    color: "#172B4D",
    boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
  },
  compactPrimaryBtn: {
    boxSizing: "border-box",
    minWidth: "64px",
    height: "22px",
    borderRadius: "14px",
    border: "1px solid rgba(0, 80, 180, 0.28)",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "4px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    cursor: "pointer",
    outline: "none",
    background: "linear-gradient(180deg, rgba(80,160,255,0.95) 0%, rgba(0,100,210,0.85) 100%)",
    color: "#FFFFFF",
    boxShadow: "0 3px 8px rgba(0,100,210,0.25), inset 0 1px 0 rgba(255,255,255,0.5)",
  },
  compactActionBtn: {
    boxSizing: "border-box",
    minWidth: "64px",
    height: "22px",
    borderRadius: "14px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "4px",
    padding: "0 8px",
    fontSize: "9px",
    fontWeight: 600,
    lineHeight: 1,
    cursor: "pointer",
    outline: "none",
    background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
    color: "#172B4D",
    boxShadow: "0 3px 8px rgba(0,0,0,0.08), inset 0 1px 0 rgba(255,255,255,1)",
  },

  pickList: {
    position: "absolute",
    left: 0,
    right: 0,
    top: "100%",
    marginTop: 6,
    background: "#fff",
    border: "1px solid #d6def2",
    borderRadius: 12,
    maxHeight: 240,
    overflow: "auto",
    zIndex: 999,
    boxShadow: "0 8px 24px rgba(0,0,0,0.08)",
  },
  pickItem: {
    width: "100%",
    textAlign: "left",
    padding: "10px 12px",
    border: "none",
    background: "transparent",
    cursor: "pointer",
    display: "flex",
    justifyContent: "space-between",
    gap: 10,
    color: "#122",
  },

  partRow: {
    display: "flex",
    gap: 10,
    alignItems: "center",
    padding: 10,
    borderRadius: 12,
    border: "1px solid #e9eefc",
    background: "#f7f9ff",
  },
  badge: {
    display: "inline-block",
    padding: "2px 8px",
    borderRadius: "16px",
    border: "1px solid rgba(255, 255, 255, 0.3)",
    background: "rgba(255, 255, 255, 0.4)",
    backdropFilter: "blur(8px)",
    color: "#42526E",
    fontSize: 10,
    marginRight: 6,
    fontWeight: 700,
  },
  threadToggle: {
    border: "1px solid rgba(255, 255, 255, 0.3)",
    background: "rgba(255, 255, 255, 0.4)",
    backdropFilter: "blur(8px)",
    borderRadius: "16px",
    padding: "2px 10px",
    fontSize: 10,
    cursor: "pointer",
    color: "#42526E",
    fontWeight: 700,
  },
  yellowGlass: {
    width: "94px",
    height: "26px",
    borderRadius: "16px",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "6px",
    fontSize: "10px",
    fontWeight: 700,
    textTransform: "uppercase",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    cursor: "default",
    transition: "all 0.2s ease",
    boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
    padding: "0 8px",
    /* Yellow Glass */
    background: "rgba(251, 191, 36, 0.4)",
    color: "#92400E",
    border: "1px solid rgba(251, 191, 36, 0.3)",
  },
};




