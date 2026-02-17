import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  getSelectedMessageContext,
  getCurrentItemToken,
  openCockpitDialog,
  subscribeToItemChanges,
  type OutlookMessageContext,
} from "../office";
import { getLinks, getOdooMeta, type LinkEntry, type OdooMeta } from "../api";
import DebugPanel from "./DebugPanel";
import AiPanel from "../ai/AiPanel";
import { SettingsPanel } from "./SettingsPanel";
import { clientLog } from "../logger";
import { getSettings } from "../settings";
import { applySkin } from "./skins";

type Tab = "odoo" | "ai" | "settings";


function encodeRecipients(list: any[] | undefined) {
  if (!list?.length) return "";
  return list
    .map((r) => `${String(r?.name || "").trim()}|${String(r?.email || "").trim()}`)
    .filter(Boolean)
    .join(";");
}

export default function App() {
  useEffect(() => {
    (async () => {
      try {
        const st = await getSettings();
        applySkin(st.skinId || "classic");
      } catch {
        applySkin("classic");
      }
    })();
  }, []);

  const [tab, setTab] = useState<Tab>("odoo");
  const [ctx, setCtx] = useState<OutlookMessageContext>({});
  const [meta, setMeta] = useState<OdooMeta | null>(null);
  const [links, setLinks] = useState<LinkEntry[]>([]);
  const [msg, setMsg] = useState<string | null>(null);
  const [showThread, setShowThread] = useState(false);

  const ctxLoadSeqRef = useRef(0);
  const lastItemTokenRef = useRef<string>("");

  const [subjectOpen, setSubjectOpen] = useState(false);
  const [summaryOpen, setSummaryOpen] = useState(false);
  const [summaryTxt, setSummaryTxt] = useState<string>("");

  const SUMMARY_KEY = "icc.summary.v2";
  const SUMMARY_KEEP_MS = 5 * 24 * 60 * 60 * 1000;

  function normStr(s: any): string {
    return String(s ?? "").trim();
  }

  function normEmail(s: any): string {
    return normStr(s).toLowerCase();
  }

  function fnv1a(str: string): string {
    let h = 2166136261;
    for (let i = 0; i < str.length; i++) {
      h ^= str.charCodeAt(i);
      // 32-bit FNV-1a
      h += (h << 1) + (h << 4) + (h << 7) + (h << 8) + (h << 24);
    }
    return (h >>> 0).toString(16);
  }

  // IMPORTANT: must match AiPanel's makeEmailKey() to keep per-email summary/workspace consistent
  function makeEmailKey(c: any): string {
    const cid = normStr(c?.conversationId);
    const imid = normStr(c?.internetMessageId);
    const itemId = normStr(c?.itemId || c?.id);

    if (cid && imid) return `${cid}::${imid}`;
    if (cid && itemId) return `${cid}::${itemId}`;

    // Fallback: stable-ish hash from visible metadata
    const subject = normStr(c?.subject).replace(/\s+/g, " ").toLowerCase();
    const from = normEmail(c?.fromEmail || c?.from);
    const to = normEmail(c?.toEmail || "");
    return `nocid::h${fnv1a([subject, from, to].join("|"))}`;
  }
  function loadSummary(emailKey: string) {
    try {
      const raw = localStorage.getItem(SUMMARY_KEY);
      const obj = raw ? (JSON.parse(raw) as Record<string, { ts: number; text: string }>) : {};
      const now = Date.now();
      // prune
      const pruned: typeof obj = {};
      for (const [k, v] of Object.entries(obj || {})) {
        if (v && typeof v.ts === "number" && typeof v.text === "string" && now - v.ts <= SUMMARY_KEEP_MS) pruned[k] = v;
      }
      if (JSON.stringify(pruned) !== JSON.stringify(obj)) localStorage.setItem(SUMMARY_KEY, JSON.stringify(pruned));
      setSummaryTxt(pruned[emailKey]?.text || "");
    } catch {
      setSummaryTxt("");
    }
  }

  const emailKey = useMemo(() => makeEmailKey(ctx as any), [ctx.conversationId, (ctx as any).internetMessageId, (ctx as any).itemId, (ctx as any).id, (ctx as any).subject, (ctx as any).fromEmail, (ctx as any).from, (ctx as any).toEmail]);

  useEffect(() => {
    setSubjectOpen(false);
    setSummaryOpen(false);
    if (emailKey) loadSummary(emailKey);
    else setSummaryTxt("");
  }, [emailKey]);

  useEffect(() => {
    function onSummaryUpdated(ev: any) {
      const k = ev?.detail?.emailKey;
      if (!k || k !== emailKey) return;

      try {
        const raw = localStorage.getItem(SUMMARY_KEY);
        const obj = raw ? (JSON.parse(raw) as Record<string, { ts: number; text: string }>) : {};
        setSummaryTxt(obj[k]?.text || "");
      } catch {
        // ignore
      }
    }

    window.addEventListener("icc-summary-updated", onSummaryUpdated);
    return () => window.removeEventListener("icc-summary-updated", onSummaryUpdated);
  }, [emailKey]);

  async function loadContextAndLinks(reason?: string) {
    const reqId = ++ctxLoadSeqRef.current;
    try {
      const c = await getSelectedMessageContext();
      if (reqId != ctxLoadSeqRef.current) return;
      setCtx(c);
      setShowThread(false);
      clientLog("info", `[taskpane] ctx updated (${reason || 'unknown'}) conversationId=${c.conversationId || ''} itemId=${(c as any).itemId || ''}`);

      if (!c.conversationId) {
        setLinks([]);
        setMsg("Office.context.mailbox.item não está disponível. Abre um email e volta a tentar.");
        return;
      }

      setMsg(null);
      try {
        const l = await getLinks(c.conversationId);
        if (reqId != ctxLoadSeqRef.current) return;
        setLinks(l);
      } catch (e: any) {
        if (reqId != ctxLoadSeqRef.current) return;
        setMsg(e?.message ?? String(e));
      }
    } catch (e: any) {
      if (reqId != ctxLoadSeqRef.current) return;
      setMsg(e?.message ?? String(e));
    }
  }

  useEffect(() => {
    loadContextAndLinks('init');
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Atualiza automaticamente quando o utilizador clica noutro email (Outlook Classic)
  useEffect(() => {
    let unsub: (() => void) | null = null;
    (async () => {
      try {
        unsub = await subscribeToItemChanges(() => {
          loadContextAndLinks();
        });
      } catch (e) {
        clientLog("warn", "[taskpane] subscribeToItemChanges failed", e);
      }
    })();
    return () => {
      try {
        unsub?.();
      } catch {
        // ignore
      }
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  // Fallback robusto: polling leve para detetar mudanca de item quando ItemChanged nao dispara/chega cedo demais
  useEffect(() => {
    let alive = true;
    let intervalId: number | null = null;

    const tick = async () => {
      try {
        const tok = await getCurrentItemToken();
        if (!alive) return;
        if (tok && tok !== lastItemTokenRef.current) {
          lastItemTokenRef.current = tok;
          loadContextAndLinks('poll-now');
          window.setTimeout(() => loadContextAndLinks('poll-late'), 450);
        }
      } catch {
        // ignore
      }
    };

    (async () => {
      try {
        lastItemTokenRef.current = await getCurrentItemToken();
      } catch {
        lastItemTokenRef.current = '';
      }
      intervalId = window.setInterval(() => {
        tick();
      }, 900);
    })();

    return () => {
      alive = false;
      if (intervalId) window.clearInterval(intervalId);
    };
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, []);

  useEffect(() => {
    (async () => {
      try {
        setMeta(await getOdooMeta());
      } catch (e: any) {
        setMsg((prev) => prev || (e?.message ?? String(e)));
      }
    })();
  }, []);

  async function refreshLinks(conversationId?: string) {
    const cid = conversationId || ctx.conversationId;
    if (!cid) return setLinks([]);
    try {
      setLinks(await getLinks(cid));
    } catch (e: any) {
      setMsg(e?.message ?? String(e));
    }
  }

  async function openDialog(targetMode: "new" | "add" | "edit", extra?: Record<string, string>) {
    if (!ctx.conversationId && targetMode !== "edit") {
      setMsg("Seleciona um email primeiro.");
      return;
    }

    try {
      await openCockpitDialog({
        mode: targetMode,
        conversationId: ctx.conversationId || "",
        internetMessageId: ctx.internetMessageId || "",
        subject: ctx.subject || "",
        fromEmail: ctx.fromEmail || "",
        fromName: ctx.fromName || "",
        receivedAtIso: ctx.receivedDateTimeIso || "",
        toR: encodeRecipients(ctx.toRecipients),
        ccR: encodeRecipients(ctx.ccRecipients),
        ...(extra || {}),
      });
      await refreshLinks();
    } catch (e: any) {
      setMsg(e?.message ?? String(e));
    }
  }

  return (
    <div style={S.shell}>
      <header style={S.header}>
        <div style={S.titleBlock}>
          <img src="/icon-32.png" alt="" style={S.titleLogo} />
          <div style={{ minWidth: 0 }}>
            <div style={S.title}>Inbox CRM Cockpit</div>
            <div style={S.subtitle}>Odoo + IA</div>
          </div>
        </div>
        <div style={S.tabs}>
          <button style={tab === "odoo" ? S.pillA : S.pill} onClick={() => setTab("odoo")} title="Integração com Odoo">
            Odoo
          </button>
          <button style={tab === "ai" ? S.pillA : S.pill} onClick={() => setTab("ai")} title="Assistente IA">
            IA
          </button>
          <button style={tab === "settings" ? S.pillA : S.pill} onClick={() => setTab("settings")} title="Definições">
            Def.
          </button>
        </div>
      </header>

      {/* Resumo (cache local 3 dias) */}
      <div style={S.slimCard}>
        <div style={S.slimTopRow}>
          <span style={S.slimLabel}>Resumo</span>
          <span style={S.flex1} />
          <button
            type="button"
            onClick={() => setSummaryOpen((v) => !v)}
            style={S.expandoBtn}
            title={summaryOpen ? "Recolher" : "Expandir"}
          >
            {summaryOpen ? "▴" : "▾"}
          </button>
        </div>

        {summaryOpen ? (
          <div style={S.summaryBody}>{summaryTxt || "—"}</div>
        ) : (
          <div style={S.summaryPreview}>
            {(summaryTxt ? (summaryTxt.split("\n")[0] || summaryTxt).slice(0, 160) : "—")}
          </div>
        )}
      </div>

      {/* Assunto / De / Thread (compacto, expansível) */}
      <div style={S.slimCard}>
        <div style={S.slimTopRow}>
          <span style={S.slimLabel}>Assunto</span>
          <span style={S.subjectLine} title={ctx.subject || ""}>
            {ctx.subject || "—"}
          </span>
          <span style={S.flex1} />
          <button
            type="button"
            onClick={() => setSubjectOpen((v) => !v)}
            style={S.expandoBtn}
            title={subjectOpen ? "Recolher detalhes" : "Expandir detalhes"}
          >
            {subjectOpen ? "▴" : "▾"}
          </button>
        </div>

        {subjectOpen && (
          <>
            <div style={S.kv}>
              <span style={S.k}>De</span>
              <span style={S.v}>{ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail}>` : ctx.fromEmail || "—"}</span>
            </div>
            <div style={S.kv}>
              <span style={S.k}>Thread</span>
              <span style={{ ...S.v, display: "flex", alignItems: "center", gap: 8 }}>
                {showThread ? (
                  <span
                    style={{
                      fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
                      fontSize: 11,
                    }}
                  >
                    {ctx.conversationId || "—"}
                  </span>
                ) : (
                  <button
                    type="button"
                    onClick={() => setShowThread(true)}
                    style={S.threadToggle}
                    title="Mostrar o ID da Thread (conversationId)"
                  >
                    Thread ▾
                  </button>
                )}

                {showThread && (
                  <button
                    type="button"
                    onClick={() => setShowThread(false)}
                    style={S.threadToggle}
                    title="Ocultar o ID da Thread"
                  >
                    ▴
                  </button>
                )}
              </span>
            </div>
          </>
        )}
      </div>

      {tab === "odoo" && (
        <>
          <div style={S.card}>
            <div style={{ display: "flex", gap: 8, flexWrap: "wrap" }}>
              <button
                style={S.btn}
                disabled={!ctx.conversationId}
                onClick={() => openDialog("new")}
                title="Criar um novo registo no Odoo e ligar ao email (abre numa janela)"
              >
                ➕ Criar
              </button>

              <button
                style={S.btn2}
                disabled={!ctx.conversationId}
                onClick={() => openDialog("add")}
                title="Ligar este email a um registo existente (abre numa janela)"
              >
                🔗 Ligar
              </button>

              <button style={S.btnGhost} onClick={() => refreshLinks()} title="Atualiza a lista de relacionados">
                ↻ Recarregar
              </button>
            </div>

            {msg && <div style={S.msg}>{msg}</div>}

            <div style={{ marginTop: 12 }}>
              <div style={S.sectionTitle}>Relacionado nesta conversa</div>
              {!links.length ? (
                <div style={S.muted}>Nada ligado ainda.</div>
              ) : (
                <div style={{ display: "grid", gap: 8, marginTop: 8 }}>
                  {links.map((l) => (
                    <div key={l.id} style={S.linkRow}>
                      <div style={{ minWidth: 0 }}>
                        <div style={S.linkTitle}>{l.title || l.model}</div>
                        <div style={S.linkMeta}>
                          {l.model} · {String(l.resId)}
                        </div>
                      </div>
                      {l.url ? (
                        <a style={S.linkA} href={l.url} target="_blank" rel="noreferrer" title="Abrir no Odoo">
                          Abrir
                        </a>
                      ) : (
                        <span style={S.linkMeta}>—</span>
                      )}
                    </div>
                  ))}
                </div>
              )}
            </div>

            <div style={{ marginTop: 10, fontSize: 11, color: "rgba(11,45,107,0.65)" }}>
              {meta ? (
                <span>
                  Odoo: {meta.baseUrl} · DB: {meta.db}
                </span>
              ) : (
                <span>Odoo: a verificar…</span>
              )}
            </div>
          </div>

          <DebugPanel />
        </>
      )}

      {tab === "ai" && <AiPanel ctx={ctx} />}

      {tab === "settings" && (
        <SettingsPanel />
      )}
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  shell: {
    padding: 10,
    background: "#eff6ff",
    minHeight: "100vh",
    fontFamily: "Segoe UI, system-ui",
    fontSize: 12,
    color: "#0b2d6b",
  },

  header: {
    maxWidth: 330,
    margin: "0 auto 10px auto",
    display: "flex",
    justifyContent: "space-between",
    alignItems: "flex-start",
    gap: 10,
    marginBottom: 10,
  },

  titleBlock: {
    display: "flex",
    alignItems: "center",
    gap: 8,
    minWidth: 0,
  },
  titleLogo: {
    width: 20,
    height: 20,
    flex: "0 0 auto",
    objectFit: "contain",
    borderRadius: 6,
  },
  title: { fontWeight: "600", fontSize: 14, lineHeight: "16px" },
  subtitle: { fontWeight: "500", fontSize: 11, color: "rgba(11,45,107,0.70)", marginTop: 2 },
  tabs: { display: "flex", gap: 6, alignItems: "center", flexWrap: "wrap" },
  pill: {
    padding: "6px 10px",
    borderRadius: "999px",
    border: "1px solid rgba(11,45,107,0.16)",
    background: "rgba(255,255,255,0.70)",
    color: "#0b2d6b",
    fontWeight: "600",
    fontSize: 12,
    cursor: "pointer",
  },
  pillA: {
    padding: "6px 10px",
    borderRadius: "999px",
    border: "1px solid rgba(11,45,107,0.20)",
    background: "#0b2d6b",
    color: "#fff",
    fontWeight: "600",
    fontSize: 12,
    cursor: "pointer",
  },

  card: {
    width: "100%",
    maxWidth: 330,
    margin: "0 auto 10px auto",
    borderRadius: "16px",
    padding: 12,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.85)",
    boxShadow: "0 1px 10px rgba(11,45,107,0.06)",
    marginBottom: 10,
    boxSizing: "border-box",
  },

  slimCard: {
    width: "100%",
    maxWidth: 330,
    margin: "0 auto 8px auto",
    borderRadius: "16px",
    padding: 10,
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.75)",
  },

  slimTopRow: {
    display: "flex",
    alignItems: "center",
    gap: 6,
  },

  slimLabel: {
    fontSize: 11,
    fontWeight: 400,
    color: "rgba(11,45,107,0.75)",
    whiteSpace: "nowrap",
  },

  flex1: { flex: 1 },

  expandoBtn: {
    borderRadius: 10,
    padding: "1px 6px",
    border: "1px solid rgba(11,45,107,0.14)",
    background: "rgba(255,255,255,0.70)",
    color: "#0b2d6b",
    fontSize: 11,
    fontWeight: 400,
    cursor: "pointer",
  },

  subjectLine: {
    fontSize: 11,
    fontWeight: 400,
    color: "#0b2d6b",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
    maxWidth: 210,
  },

  summaryPreview: {
    marginTop: 6,
    fontSize: 9,
    lineHeight: "12px",
    color: "rgba(11,45,107,0.85)",
    display: "-webkit-box",
    WebkitLineClamp: 6,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
    whiteSpace: "pre-wrap",
  },

  summaryBody: {
    marginTop: 8,
    fontSize: 9,
    lineHeight: "12px",
    color: "rgba(11,45,107,0.90)",
    whiteSpace: "pre-wrap",
  },

  kv: { display: "grid", gridTemplateColumns: "70px 1fr", gap: 8, alignItems: "start", marginBottom: 6 },
  k: { fontWeight: "600", color: "rgba(11,45,107,0.75)" },
  v: { fontWeight: "500", color: "#0b2d6b", overflow: "hidden", textOverflow: "ellipsis" },

  sectionTitle: { fontWeight: "600", fontSize: 12, marginTop: 2 },
  muted: { fontSize: 12, color: "rgba(11,45,107,0.65)", marginTop: 6 },

  btn: {
    borderRadius: "12px",
    padding: "8px 10px",
    border: "1px solid rgba(11,45,107,0.20)",
    background: "#0b2d6b",
    color: "#fff",
    fontWeight: "600",
    cursor: "pointer",
  },
  btn2: {
    borderRadius: "12px",
    padding: "8px 10px",
    border: "1px solid rgba(11,45,107,0.20)",
    background: "rgba(11,45,107,0.10)",
    color: "#0b2d6b",
    fontWeight: "600",
    cursor: "pointer",
  },
  btnGhost: {
    borderRadius: "12px",
    padding: "8px 10px",
    border: "1px solid rgba(11,45,107,0.16)",
    background: "rgba(255,255,255,0.70)",
    color: "#0b2d6b",
    fontWeight: "600",
    cursor: "pointer",
  },

  msg: {
    marginTop: 10,
    borderRadius: "12px",
    padding: 10,
    border: "1px solid rgba(245, 158, 11, 0.45)",
    background: "rgba(245, 158, 11, 0.10)",
    color: "#7a4a00",
    fontSize: 12,
  },

  linkRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: 10,
    padding: 10,
    borderRadius: "12px",
    border: "1px solid rgba(11,45,107,0.12)",
    background: "rgba(255,255,255,0.70)",
  },
  linkTitle: { fontWeight: "600", fontSize: 12, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  linkMeta: { fontSize: 11, color: "rgba(11,45,107,0.65)" },
  linkA: { fontSize: 11, fontWeight: "600", color: "#0b2d6b", textDecoration: "none" },
};
