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
  findOdooField,
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
function AiAssistant({ bodyText, onAddAction }: { bodyText: string, onAddAction: (title: string, type: string) => void }) {
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
    // Auto-trigger apenas se tivermos corpo e ainda não tivermos analisado
    if (bodyText && !data && !loading) {
      analyze();
    }
  }, [bodyText]);

  if (!data || !data.summary.length) {
    if (loading) return (
      <div style={{ marginBottom: 16, padding: 12, borderRadius: 12, border: "1px dashed #d6def2", background: "rgba(255,255,255,0.4)" }}>
        <div style={{ fontSize: 10, fontWeight: 800, color: "#2563eb", marginBottom: 4, display: "flex", alignItems: "center", gap: "6px" }}>
          <Icons.RefreshCw size={10} className="animate-spin" />
          ASSISTENTE IA • A ANALISAR...
        </div>
      </div>
    );

    return (
      <div style={{ marginBottom: 16, padding: 12, borderRadius: 12, border: "1px solid #d6def2", background: "rgba(255,255,255,0.4)", backdropFilter: "blur(12px)" }}>
        <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", marginBottom: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
            <Icons.Sparkles size={12} color="#2563eb" />
            <div style={{ fontSize: 10, fontWeight: 800, color: "#2563eb", textTransform: "uppercase" }}>Assistente IA</div>
          </div>
          <button style={S.secondaryBtn} onClick={analyze} title="Reanalisar o conteúdo">
            <Icons.RefreshCw size={10} />
            REANALISAR
          </button>
        </div>
        <div style={{ fontSize: 11, color: bodyText ? "#BF2600" : "#777" }}>
          {bodyText ? "⚠️ Clique em Reanalisar para processar o email." : "ℹ️ O conteúdo do email ainda não foi carregado."}
        </div>
        {error && <div style={{ fontSize: 10, color: "#BF2600", marginTop: 8 }}><b>ERRO:</b> {error}</div>}
      </div>
    );
  }

  return (
    <div style={{ marginBottom: 16, padding: 12, borderRadius: 12, border: "1px solid #d6def2", background: "rgba(255,255,255,0.4)", backdropFilter: "blur(12px)" }}>
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
              style={S.primaryBtn}
              onClick={() => onAddAction(act, "project.task")}
              title="Criar tarefa"
            >
              <Icons.Plus size={10} />
              {act.length > 12 ? act.substring(0, 11) + ".." : act.toUpperCase()}
            </button>
          ))}
        </div>
      )}
    </div>
  );
}
type Mode = "new" | "add" | "edit";
type Entity = "project.task" | "helpdesk.ticket" | "project.project" | "crm.lead" | "res.partner";

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

  return (
    <div style={{ marginTop: 16 }}>
      <label style={S.labBlock}>ANEXOS ({attachments.length})</label>
      <div style={{ display: "flex", flexWrap: "wrap", gap: "6px", marginTop: "4px" }}>
        {attachments.map(att => {
          const isSelected = selected.includes(att.name);
          return (
            <button
              key={att.name}
              onClick={() => onToggle(att.name)}
              style={isSelected ? S.primaryBtn : S.secondaryBtn}
              title={att.name}
            >
              {isSelected ? "✅" : "📎"} {att.name.length > 8 ? att.name.substring(0, 7) + ".." : att.name}
            </button>
          );
        })}
      </div>
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

type Ctx = {
  conversationId: string;
  internetMessageId: string;
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
    <div style={{ marginTop: 10, position: "relative" }}>
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
            className="jira-ghost-button" style={S.btn2}
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

      {pickedId ? (
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

export default function DialogApp() {
  const isDevRuntime = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
  const [mode, setMode] = useState<Mode>(() => getMode());
  const [editId, setEditId] = useState<string | null>(() => qp().get("recordId") || null);
  const [ctx, setCtx] = useState<Ctx>(() => getCtxFromQuery());
  const [showThread, setShowThread] = useState(false);
  const [entity, setEntity] = useState<Entity>(() => {
    const m = qp().get("model") || "";
    return (m as Entity) || "project.task";
  });
  const [status, setStatus] = useState<string | null>(null);

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
    const a = localStorage.getItem("ic_bridge_atts");
    if (b) setFullBody(b);
    if (a) {
      try { setEmailAtts(JSON.parse(a)); } catch { }
    }
  }, []);

  useEffect(() => {
    (async () => {
      try {
        const st = await getSettings();
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

          {mode === "add" ? (
            <AddExistingPanel entity={entity} ctx={ctx} onStatus={setStatus} />
          ) : entity === "project.task" && (
            <TaskForm
              mode={mode}
              ctx={ctx}
              editId={editId}
              fullBody={fullBody}
              emailAtts={emailAtts}
              onStatus={setStatus}
              fromEmail={ctx.fromEmail}
            />
          )}
          {entity === "project.project" && (
            <ProjectForm
              mode={mode}
              ctx={ctx}
              editId={editId}
              fullBody={fullBody}
              emailAtts={emailAtts}
              onStatus={setStatus}
              fromEmail={ctx.fromEmail}
            />
          )}
          {entity === "crm.lead" && (
            <LeadForm
              mode={mode}
              ctx={ctx}
              editId={editId}
              fullBody={fullBody}
              emailAtts={emailAtts}
              onStatus={setStatus}
              fromEmail={ctx.fromEmail}
            />
          )}
          {entity === "res.partner" && <ContactHubForm mode={mode} ctx={ctx} editId={editId} onStatus={setStatus} />}
          {entity === "helpdesk.ticket" && (
            <HelpdeskTicketForm
              mode={mode}
              ctx={ctx}
              editId={editId}
              fullBody={fullBody}
              emailAtts={emailAtts}
              onStatus={setStatus}
            />
          )}

          {status && <div style={S.alert}>{status}</div>}
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

  useEffect(() => {
    if (mode === "new" && fullBody) setDescription(fullBody);
  }, [mode, fullBody]);

  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);

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
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("project.task", [editId], ["name", "description", "project_id", "user_ids", "date_deadline", "stage_id", "parent_id"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setDescription(r.description || "");
        if (r.project_id) { setProjectId(r.project_id[0]); setProjectName(r.project_id[1]); }
        if (Array.isArray(r.user_ids) && r.user_ids.length) {
          const u = await readOdoo("res.users", [r.user_ids[0]], ["name"]);
          setAssigneeId(r.user_ids[0]);
          setAssigneeName(u?.[0]?.name || "");
        }
        if (r.date_deadline) setDeadline(String(r.date_deadline));
        if (r.stage_id) { setStageId(r.stage_id[0]); setStageName(r.stage_id[1]); }
        if (r.parent_id) { setIsSub(true); setParentId(r.parent_id[0]); setParentName(r.parent_id[1]); }
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [mode, editId]);

  useEffect(() => {
    (async () => {
      try {
        if (!projectId) return setStagePick([]);
        const rows = await searchOdooDomain(
          "project.task.type",
          ["|", ["project_ids", "=", false], ["project_ids", "in", [projectId]]],
          ["id", "name"],
          50
        );
        setStagePick(rows || []);
      } catch {
        setStagePick([]);
      }
    })();
  }, [projectId]);

  async function save() {
    try {
      let values: any = {
        name: name || "Nova tarefa",
        description: description || "",
      };
      if (projectId) values.project_id = projectId;
      if (assigneeId) values.user_ids = [[6, 0, [assigneeId]]];
      if (deadline) values.date_deadline = deadline;
      if (stageId) values.stage_id = stageId;
      if (isSub && parentId) values.parent_id = parentId;

      let id = editId;

      if (mode === "edit") {
        await writeOdoo("project.task", id, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("project.task", values);
      id = await createOdoo("project.task", values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "project.task",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Tarefa criada + Ligada ✅");

      // Criar subtarefas pendentes
      if (pendingSubtasks.length > 0) {
        onStatus(`A criar ${pendingSubtasks.length} subtarefas detetadas...`);
        for (const subTitle of pendingSubtasks) {
          try {
            const subtaskValues = await withReferenceCode("project.task", {
              name: subTitle,
              project_id: projectId || false,
              parent_id: id,
              user_ids: assigneeId ? [assigneeId] : [],
            });
            await createOdoo("project.task", subtaskValues);
          } catch (e) {
            console.error("Erro ao criar subtarefa diferida", e);
          }
        }
      }

      onStatus("Criado ✅");

      // Handle Attachments
      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        for (const name of selectedAtts) {
          const att = (emailAtts || []).find((a: any) => a.name === name);
          if (att) {
            try {
              await createOdoo("ir.attachment", {
                name: att.name,
                datas: att.content,
                datas_fname: att.name,
                mimetype: att.contentType,
                res_model: "project.task",
                res_id: id,
                type: "binary"
              });
            } catch (e) {
              console.error("Erro ao enviar anexo", att.name, e);
              // keep going with other attachments
            }
          }
        }
      }

      onStatus("Sucesso! ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function handleAddAiAction(title: string) {
    if (mode === "new") {
      setPendingSubtasks(prev => [...prev, title]);
      onStatus(`Subtarefa "${title}" agendada para criação ✅`);
      return;
    }

    onStatus(`A criar tarefa: ${title}...`);
    try {
      let val: any = {
        name: title,
        project_id: projectId || false,
        user_ids: assigneeId ? [assigneeId] : [],
      };
      if (mode === "edit" && editId) val.parent_id = editId;

      val = await withReferenceCode("project.task", val);
      const newId = await createOdoo("project.task", val);
      onStatus(`Tarefa "${title}" criada (#${newId}) ✅`);
    } catch (e: any) {
      onStatus(`Falha ao criar: ${e.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck projectId={projectId} fromEmail={fromEmail} />

      <div style={S.row}>
        <label style={S.lab}>TÍTULO</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Título da tarefa" />
      </div>

      <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} />

      {pendingSubtasks.length > 0 && (
        <div style={{ marginTop: 8, padding: "4px 8px", background: "#E3F2FD", borderRadius: 8, fontSize: 10, color: "#0D47A1", fontWeight: 700 }}>
          AGENDADAS: {pendingSubtasks.map(s => `"${s}"`).join(", ")}
        </div>
      )}

      <TypeaheadPicker
        label="PROJETO"
        placeholder="Pesquisar projeto…"
        model="project.project"
        pickedId={projectId}
        pickedName={projectName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setProjectId(id);
          setProjectName(id ? (it.display_name || it.name || `#${id}`) : "");
          if (!id) { setStageId(null); setStageName(""); }
        }}
      />

      <TypeaheadPicker
        label="RESPONSÁVEL"
        placeholder="Pesquisar utilizador…"
        model="res.users"
        fields={["id", "name", "display_name", "email"]}
        pickedId={assigneeId}
        pickedName={assigneeName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setAssigneeId(id);
          setAssigneeName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={S.grid2}>
        <PickerStatic
          label="ETAPA"
          pickedId={stageId}
          pickedName={stageName}
          items={stagePick}
          onPick={(it: any) => { setStageId(it.id); setStageName(it.name || it.display_name || `#${it.id}`); }}
          placeholder={projectId ? "Escolher etapa..." : "Etapa (opcional)"}
        />

        <div style={S.row}>
          <label style={S.lab}>PRAZO</label>
          <input style={S.input} type="date" value={deadline} onChange={(e) => setDeadline(e.target.value)} />
        </div>
      </div>

      <div style={S.row}>
        <label style={S.lab}>SUBTAREFA</label>
        <input
          type="checkbox"
          checked={isSub}
          onChange={(e) => {
            setIsSub(e.target.checked);
            if (!e.target.checked) { setParentId(null); setParentName(""); }
          }}
        />
      </div>

      {isSub ? (
        <TypeaheadPicker
          label="PARENT TASK"
          placeholder={projectId ? "Pesquisar tarefa (filtra por projeto)..." : "Pesquisar tarefa (global)..."}
          model="project.task"
          fields={["id", "name", "display_name", "project_id"]}
          pickedId={parentId}
          pickedName={parentName}
          extraDomain={(q) => {
            const d: any[] = [];
            if (projectId) d.push(["project_id", "=", projectId]);
            if (q?.trim()) d.push(["name", "ilike", q.trim()]);
            return d;
          }}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setParentId(id);
            setParentName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />
      ) : null}


      <div style={{ marginTop: 12 }}>
        <label style={S.labBlock}>DESCRIÇÃO</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Descrição / notas..." />
      </div>

      <AttachmentPicker
        attachments={emailAtts}
        selected={selectedAtts}
        onToggle={(name) => setSelectedAtts(prev => prev.includes(name) ? prev.filter(n => n !== name) : [...prev, name])}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button>
      </div>
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

  useEffect(() => {
    if (mode === "new" && (ctx.bodyHtml || fullBody)) setDescription(ctx.bodyHtml || fullBody);
  }, [mode, ctx.bodyHtml, fullBody]);

  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);

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
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        if (r.partner_id) {
          setPartnerId(r.partner_id[0]);
          setPartnerName(r.partner_id[1]);
        }
        if (r.user_id) {
          setManagerId(r.user_id[0]);
          setManagerName(r.user_id[1]);
        }
        if (r.description) setDescription(String(r.description));
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
  }, [mode, editId]);

  async function save() {
    try {
      let values: any = { name: name || "Novo projeto" };
      if (partnerId) values.partner_id = partnerId;
      if (managerId) values.user_id = managerId;
      if (description) values.description = description;
      if (ctx.bodyHtml) values.bodyHtml = ctx.bodyHtml; // Server can use this for the first message post too if needed

      if (mode === "edit") {
        try {
          await writeOdoo("project.project", editId, values);
        } catch {
          const v2 = { ...values };
          delete v2.description;
          await writeOdoo("project.project", editId, v2);
        }
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("project.project", values);
      let id = await createOdoo("project.project", values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "project.project",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        for (const fname of selectedAtts) {
          const att = (emailAtts || []).find((a: any) => a.name === fname);
          if (att) {
            await createOdoo("ir.attachment", {
              name: att.name,
              datas: att.content,
              res_model: "project.project",
              res_id: id,
              type: "binary"
            });
          }
        }
      }
      onStatus("Sucesso! ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function handleAddAiAction(title: string) {
    onStatus(`A criar tarefa IA: ${title}...`);
    try {
      let val: any = {
        name: title,
        project_id: editId || false,
        partner_id: partnerId || false,
      };
      val = await withReferenceCode("project.task", val);
      const nId = await createOdoo("project.task", val);
      onStatus(`Tarefa IA detetada e criada (#${nId}) ✅`);
    } catch (e: any) {
      onStatus(`Falha IA: ${e.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck partnerId={partnerId} fromEmail={fromEmail} />

      <div style={S.row}>
        <label style={S.lab}>NOME</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do projeto" />
      </div>

      <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} />

      <TypeaheadPicker
        label="CLIENTE"
        placeholder="Pesquisar contacto/empresa…"
        model="res.partner"
        pickedId={partnerId}
        pickedName={partnerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPartnerId(id);
          setPartnerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <TypeaheadPicker
        label="GESTOR"
        placeholder="Pesquisar utilizador…"
        model="res.users"
        fields={["id", "name", "display_name", "email"]}
        pickedId={managerId}
        pickedName={managerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setManagerId(id);
          setManagerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ marginTop: 12 }}>
        <label style={S.labBlock}>DESCRIÇÃO</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Notas do projeto…" />
      </div>

      <AttachmentPicker
        attachments={emailAtts}
        selected={selectedAtts}
        onToggle={(fname) => setSelectedAtts(prev => prev.includes(fname) ? prev.filter(n => n !== fname) : [...prev, fname])}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button style={S.btn} onClick={save}>
          {mode === "edit" ? "Guardar" : "Criar"}
        </button>
      </div>
    </div>
  );
}

function LeadForm({ mode, ctx, editId, onStatus, fullBody, emailAtts, fromEmail }: any) {
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
  const [leadTypeValue, setLeadTypeValue] = useState("");
  const [leadTypeRelationId, setLeadTypeRelationId] = useState<number | null>(null);
  const [leadTypeRelationName, setLeadTypeRelationName] = useState("");

  async function resolveLeadTypeField(): Promise<OdooFieldMeta | null> {
    const field = await findOdooField("crm.lead", {
      labels: ["Tipo de Lead", "Lead Type", "Tipo Lead", "Tipo da Lead"],
      nameCandidates: [
        "x_studio_tipo_de_lead",
        "x_studio_tipo_lead",
        "x_tipo_de_lead",
        "x_tipo_lead",
        "lead_type",
        "x_lead_type",
        "tipo_de_lead",
        "tipo_lead",
      ],
      namePatterns: [/tipo.*lead/i, /lead.*tipo/i, /lead_type/i, /x_studio.*tipo/i, /x_studio.*lead.*type/i],
      preferredTypes: ["selection", "many2one"],
    });
    setLeadTypeField(field);
    return field;
  }

  useEffect(() => {
    if (mode === "new" && (ctx.bodyHtml || fullBody)) setDescription(ctx.bodyHtml || fullBody);
  }, [mode, ctx.bodyHtml, fullBody]);

  const [selectedAtts, setSelectedAtts] = useState<string[]>([]);

  useEffect(() => {
    let alive = true;
    (async () => {
      try {
        const field = await resolveLeadTypeField();
        if (!alive) return;
        setLeadTypeField(field);
      } catch {
        if (alive) setLeadTypeField(null);
      } finally {
        if (alive) setLeadTypeLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        let rows: any[] | null = null;
        const fieldName = leadTypeField?.name;
        const baseFields = ["name", "contact_name", "email_from", "phone", "partner_id", "stage_id"];
        try {
          rows = await readOdoo("crm.lead", [editId], fieldName ? [...baseFields, fieldName, "description"] : [...baseFields, "description"]);
        } catch {
          rows = await readOdoo("crm.lead", [editId], fieldName ? [...baseFields, fieldName] : baseFields);
        }
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setContactName(r.contact_name || "");
        setEmail(r.email_from || "");
        setPhone(r.phone || "");
        if (r.partner_id) { setPartnerId(r.partner_id[0]); setPartnerName(r.partner_id[1]); }
        if (r.stage_id) { setStageId(r.stage_id[0]); setStageName(r.stage_id[1]); }
        if (r.description) setDescription(String(r.description));
        if (fieldName) {
          const fieldValue = r[fieldName];
          if (leadTypeField?.type === "many2one") {
            if (Array.isArray(fieldValue) && fieldValue[0]) {
              setLeadTypeRelationId(fieldValue[0]);
              setLeadTypeRelationName(fieldValue[1] || `#${fieldValue[0]}`);
            } else {
              setLeadTypeRelationId(null);
              setLeadTypeRelationName("");
            }
          } else {
            setLeadTypeValue(String(fieldValue || ""));
          }
        } else {
          setLeadTypeValue("");
          setLeadTypeRelationId(null);
          setLeadTypeRelationName("");
        }
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
  }, [mode, editId, leadTypeField]);

  async function save() {
    try {
      const effectiveLeadTypeField = leadTypeField || await resolveLeadTypeField();
      let values: any = {
        name: name || `Lead: ${ctx.subject || "sem assunto"}`,
        contact_name: contactName || "",
        email_from: email || "",
      };
      if (phone) values.phone = phone;
      if (partnerId) values.partner_id = partnerId;
      if (stageId) values.stage_id = stageId;
      if (description) values.description = description;
      if (ctx.bodyHtml) values.bodyHtml = ctx.bodyHtml;
      if (effectiveLeadTypeField?.name) {
        if (effectiveLeadTypeField.type === "many2one") values[effectiveLeadTypeField.name] = leadTypeRelationId || false;
        else values[effectiveLeadTypeField.name] = leadTypeValue || false;
      }

      if (mode === "edit") {
        try {
          await writeOdoo("crm.lead", editId, values);
        } catch {
          const v2 = { ...values };
          delete v2.description;
          await writeOdoo("crm.lead", editId, v2);
        }
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("crm.lead", values);
      let id = await createOdoo("crm.lead", values);

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "crm.lead",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      onStatus("Criado + Ligado ✅");

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        for (const fname of selectedAtts) {
          const att = (emailAtts || []).find((a: any) => a.name === fname);
          if (att) {
            await createOdoo("ir.attachment", {
              name: att.name,
              datas: att.content,
              res_model: "crm.lead",
              res_id: id,
              type: "binary"
            });
          }
        }
      }
      onStatus("Sucesso! ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  async function handleAddAiAction(title: string) {
    onStatus(`A criar tarefa IA: ${title}...`);
    try {
      let val: any = {
        name: title,
        partner_id: partnerId || false,
      };
      val = await withReferenceCode("project.task", val);
      const nId = await createOdoo("project.task", val);
      onStatus(`Tarefa IA detetada e criada (#${nId}) ✅`);
    } catch (e: any) {
      onStatus(`Falha IA: ${e.message}`);
    }
  }

  return (
    <div>
      <OdooMemoryCheck partnerId={partnerId} fromEmail={fromEmail} />

      <div style={S.row}>
        <label style={S.lab}>NOME LEAD</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do lead" />
      </div>

      <AiAssistant bodyText={fullBody} onAddAction={handleAddAiAction} />

      <div style={S.row}>
        <label style={S.lab}>CONTACTO</label>
        <input style={S.input} value={contactName} onChange={(e) => setContactName(e.target.value)} placeholder="Nome do contacto" />
      </div>

      <div style={S.row}>
        <label style={S.lab}>EMAIL</label>
        <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
      </div>

      <div style={S.row}>
        <label style={S.lab}>TELEFONE</label>
        <input style={S.input} value={phone} onChange={(e) => setPhone(e.target.value)} placeholder="Telefone" />
      </div>

      {leadTypeLoading && (
        <div style={S.row}>
          <label style={S.lab}>TIPO DE LEAD</label>
          <input style={S.input} value={leadTypeValue} onChange={(e) => setLeadTypeValue(e.target.value)} placeholder="A carregar tipo de lead..." />
        </div>
      )}

      {!leadTypeLoading && leadTypeField?.type === "selection" && (
        <div style={S.row}>
          <label style={S.lab}>{(leadTypeField.string || "Tipo de Lead").toUpperCase()}</label>
          <select style={S.sel} value={leadTypeValue} onChange={(e) => setLeadTypeValue(e.target.value)}>
            <option value="">Selecionar...</option>
            {(leadTypeField.selection || []).map(([value, label]) => (
              <option key={value} value={value}>{label}</option>
            ))}
          </select>
        </div>
      )}

      {!leadTypeLoading && leadTypeField?.type === "many2one" && leadTypeField.relation && (
        <TypeaheadPicker
          label={(leadTypeField.string || "Tipo de Lead").toUpperCase()}
          placeholder="Pesquisar tipo de lead..."
          model={leadTypeField.relation}
          fields={["id", "name", "display_name"]}
          pickedId={leadTypeRelationId}
          pickedName={leadTypeRelationName}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setLeadTypeRelationId(id);
            setLeadTypeRelationName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />
      )}

      {!leadTypeLoading && leadTypeField && leadTypeField.type !== "selection" && leadTypeField.type !== "many2one" && (
        <div style={S.row}>
          <label style={S.lab}>{(leadTypeField.string || "Tipo de Lead").toUpperCase()}</label>
          <input
            style={S.input}
            value={leadTypeValue}
            onChange={(e) => setLeadTypeValue(e.target.value)}
            placeholder="Tipo de lead"
          />
        </div>
      )}

      {!leadTypeLoading && !leadTypeField && (
        <div style={S.row}>
          <label style={S.lab}>TIPO DE LEAD</label>
          <input
            style={S.input}
            value={leadTypeValue}
            onChange={(e) => setLeadTypeValue(e.target.value)}
            placeholder="Tipo de lead"
          />
        </div>
      )}

      <TypeaheadPicker
        label="EMPRESA"
        placeholder="Pesquisar res.partner…"
        model="res.partner"
        pickedId={partnerId}
        pickedName={partnerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPartnerId(id);
          setPartnerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <TypeaheadPicker
        label="ETAPA"
        placeholder="Pesquisar etapa do lead…"
        model="crm.stage"
        fields={["id", "name"]}
        pickedId={stageId}
        pickedName={stageName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setStageId(id);
          setStageName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={{ marginTop: 12 }}>
        <label style={S.labBlock}>DESCRIÇÃO</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Notas do lead…" />
      </div>

      <AttachmentPicker
        attachments={emailAtts}
        selected={selectedAtts}
        onToggle={(fname) => setSelectedAtts(prev => prev.includes(fname) ? prev.filter(n => n !== fname) : [...prev, fname])}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button style={S.btn} onClick={save}>
          {mode === "edit" ? "Guardar" : "Criar"}
        </button>
      </div>
    </div>
  );
}

function ContactHubForm({ mode, ctx, editId, onStatus }: any) {
  const [name, setName] = useState(ctx.fromName || ctx.subject || "");
  const [email, setEmail] = useState(ctx.fromEmail || "");
  const [phone, setPhone] = useState("");

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
        const rows = await readOdoo("res.partner", [editId], ["name", "email", "phone"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setEmail(r.email || "");
        setPhone(r.phone || "");
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

  async function saveMain() {
    try {
      if (mode === "edit") {
        await writeOdoo("res.partner", editId, { name: name || email || "Contacto", email, phone });
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      const id = await createOdoo("res.partner", { name: name || email || "Contacto", email, phone });
      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "res.partner",
        recordId: id,
        recordName: name || email,
        internetMessageId: ctx.internetMessageId,
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
        <label style={S.lab}>NOME</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Nome do contacto" />
      </div>

      <div style={S.row}>
        <label style={S.lab}>EMAIL</label>
        <input style={S.input} value={email} onChange={(e) => setEmail(e.target.value)} placeholder="email@..." />
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

  useEffect(() => {
    if (mode === "new" && (ctx.bodyHtml || fullBody)) setDescription(ctx.bodyHtml || fullBody);
  }, [mode, ctx.bodyHtml, fullBody]);

  useEffect(() => {
    if (mode !== "edit" || !editId) return;
    (async () => {
      try {
        const rows = await readOdoo("helpdesk.ticket", [editId], ["name", "description", "partner_id", "team_id", "user_id", "stage_id", "priority"]);
        const r = rows?.[0];
        if (!r) return;
        setName(r.name || "");
        setDescription(String(r.description || ""));
        if (r.partner_id) { setPartnerId(r.partner_id[0]); setPartnerName(r.partner_id[1]); }
        if (r.team_id) { setTeamId(r.team_id[0]); setTeamName(r.team_id[1]); }
        if (r.user_id) { setAssigneeId(r.user_id[0]); setAssigneeName(r.user_id[1]); }
        if (r.stage_id) { setStageId(r.stage_id[0]); setStageName(r.stage_id[1]); }
        setPriority(String(r.priority ?? "0"));
      } catch (e: any) {
        onStatus(e?.message ?? String(e));
      }
    })();
  }, [mode, editId]);

  async function save() {
    try {
      let values: any = {
        name: name || `Ticket: ${ctx.subject || "sem assunto"}`,
      };
      if (description) values.description = description;
      if (partnerId) values.partner_id = partnerId;
      if (teamId) values.team_id = teamId;
      if (assigneeId) values.user_id = assigneeId;
      if (stageId) values.stage_id = stageId;
      if (priority) values.priority = priority;

      let id = editId;
      if (mode === "edit") {
        await writeOdoo("helpdesk.ticket", id, values);
        onStatus("Atualizado ✅");
        setTimeout(() => closeDialog(), 500);
        return;
      }

      values = await withReferenceCode("helpdesk.ticket", values);
      id = await createOdoo("helpdesk.ticket", values);
      onStatus("Ticket criado ✅");

      await linkEmailToRecord({
        conversationId: ctx.conversationId,
        model: "helpdesk.ticket",
        recordId: id,
        recordName: values.name,
        internetMessageId: ctx.internetMessageId,
        subject: ctx.subject,
        fromEmail: ctx.fromEmail,
        fromName: ctx.fromName,
        receivedAtIso: ctx.receivedAtIso,
        emailWebLink: ctx.emailWebLink,
      });

      if (selectedAtts.length > 0) {
        onStatus("A enviar anexos...");
        for (const fname of selectedAtts) {
          const att = (emailAtts || []).find((a: any) => a.name === fname);
          if (!att) continue;
          await createOdoo("ir.attachment", {
            name: att.name,
            datas: att.content,
            datas_fname: att.name,
            mimetype: att.contentType,
            res_model: "helpdesk.ticket",
            res_id: id,
            type: "binary",
          });
        }
      }

      onStatus("Sucesso! ✅");
      setTimeout(() => closeDialog(), 500);
    } catch (e: any) {
      onStatus(e?.message ?? String(e));
    }
  }

  return (
    <div>
      <div style={S.row}>
        <label style={S.lab}>TÍTULO</label>
        <input style={S.input} value={name} onChange={(e) => setName(e.target.value)} placeholder="Título do ticket" />
      </div>

      <TypeaheadPicker
        label="CONTACTO"
        placeholder="Pesquisar res.partner…"
        model="res.partner"
        pickedId={partnerId}
        pickedName={partnerName}
        onPick={(it: any) => {
          const id = it?.id ?? null;
          setPartnerId(id);
          setPartnerName(id ? (it.display_name || it.name || `#${id}`) : "");
        }}
      />

      <div style={S.grid2}>
        <TypeaheadPicker
          label="EQUIPA"
          placeholder="Pesquisar equipa…"
          model="helpdesk.team"
          fields={["id", "name"]}
          pickedId={teamId}
          pickedName={teamName}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setTeamId(id);
            setTeamName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />

        <TypeaheadPicker
          label="RESPONSÁVEL"
          placeholder="Pesquisar utilizador…"
          model="res.users"
          fields={["id", "name", "display_name"]}
          pickedId={assigneeId}
          pickedName={assigneeName}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setAssigneeId(id);
            setAssigneeName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />
      </div>

      <div style={S.grid2}>
        <TypeaheadPicker
          label="ETAPA"
          placeholder="Pesquisar etapa do ticket…"
          model="helpdesk.stage"
          fields={["id", "name"]}
          pickedId={stageId}
          pickedName={stageName}
          onPick={(it: any) => {
            const id = it?.id ?? null;
            setStageId(id);
            setStageName(id ? (it.display_name || it.name || `#${id}`) : "");
          }}
        />

        <div style={S.row}>
          <label style={S.lab}>PRIORIDADE</label>
          <select style={S.sel} value={priority} onChange={(e) => setPriority(e.target.value)}>
            <option value="0">Baixa</option>
            <option value="1">Média</option>
            <option value="2">Alta</option>
            <option value="3">Urgente</option>
          </select>
        </div>
      </div>

      <div style={{ marginTop: 12 }}>
        <label style={S.labBlock}>DESCRIÇÃO</label>
        <textarea style={S.ta} value={description} onChange={(e) => setDescription(e.target.value)} placeholder="Detalhes do ticket…" />
      </div>

      <AttachmentPicker
        attachments={emailAtts}
        selected={selectedAtts}
        onToggle={(fname) => setSelectedAtts(prev => prev.includes(fname) ? prev.filter(n => n !== fname) : [...prev, fname])}
      />

      <div style={{ display: "flex", gap: 10, marginTop: 16 }}>
        <button style={S.btn} onClick={save}>{mode === "edit" ? "Guardar" : "Criar"}</button>
      </div>
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
  ta: { width: "100%", minHeight: 80, padding: "8px", border: "2px solid #DFE1E6", borderRadius: 3, resize: "vertical", color: "#172B4D", background: "#FAFBFC", fontSize: 13, outline: "none" },

  grid2: { display: "grid", gridTemplateColumns: "1fr 1fr", gap: 12 },

  btn: {
    boxSizing: "border-box",
    width: "94px", minWidth: "94px", maxWidth: "94px",
    height: "26px", minHeight: "26px", maxHeight: "26px",
    borderRadius: "16px",
    border: "1px solid rgba(0, 80, 180, 0.4)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "10px",
    fontWeight: 800,
    lineHeight: 1,
    textTransform: "uppercase",
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
    width: "94px", minWidth: "94px", maxWidth: "94px",
    height: "26px", minHeight: "26px", maxHeight: "26px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-start",
    gap: "5px",
    padding: "0 8px",
    fontSize: "10px",
    fontWeight: 800,
    lineHeight: 1,
    textTransform: "uppercase",
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
    width: "94px", minWidth: "94px", maxWidth: "94px",
    height: "26px", minHeight: "26px", maxHeight: "26px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-start",
    gap: "5px",
    padding: "0 8px",
    fontSize: "10px",
    fontWeight: 800,
    lineHeight: 1,
    textTransform: "uppercase",
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
    width: "94px", minWidth: "94px", maxWidth: "94px",
    height: "26px", minHeight: "26px", maxHeight: "26px",
    borderRadius: "16px",
    border: "1px solid rgba(0, 80, 180, 0.4)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "center",
    gap: "5px",
    padding: "0 8px",
    fontSize: "10px",
    fontWeight: 800,
    lineHeight: 1,
    textTransform: "uppercase",
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
    width: "94px", minWidth: "94px", maxWidth: "94px",
    height: "26px", minHeight: "26px", maxHeight: "26px",
    borderRadius: "16px",
    border: "1px solid rgba(200, 210, 230, 0.6)",
    backdropFilter: "blur(12px)",
    WebkitBackdropFilter: "blur(12px)",
    display: "flex",
    alignItems: "center",
    justifyContent: "flex-start",
    gap: "5px",
    padding: "0 8px",
    fontSize: "10px",
    fontWeight: 800,
    lineHeight: 1,
    textTransform: "uppercase",
    cursor: "pointer",
    flexShrink: 0,
    margin: 0,
    outline: "none",
    background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
    color: "#172B4D",
    boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
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
