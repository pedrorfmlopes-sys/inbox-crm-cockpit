import React, { useEffect, useMemo, useState } from "react";
import { getSelectedMessageContext, openCockpitDialog, type OutlookMessageContext } from "../office";
import { getLinks, getOdooMeta, type LinkEntry, type OdooMeta } from "../api";

type Tab = "odoo" | "ai" | "settings";

function EyeIcon({ off }: { off?: boolean }) {
  // Ícone pequeno, sem libs externas (14px). "off" = olho cortado.
  return off ? (
    <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" style={{ display: "block" }}>
      <path
        fill="currentColor"
        d="M2.1 3.51 3.51 2.1 21.9 20.49 20.49 21.9l-3.06-3.06A11.6 11.6 0 0 1 12 20C6.5 20 2.1 16.4 0.5 12c.73-1.96 2.02-3.73 3.69-5.15L2.1 3.51Zm6.28 6.28A3 3 0 0 0 12 15c.44 0 .85-.08 1.24-.22l-4.86-4.86ZM12 4c5.5 0 9.9 3.6 11.5 8a13.2 13.2 0 0 1-4.1 5.52l-1.44-1.44A11 11 0 0 0 21.5 12C20.1 8.3 16.4 6 12 6c-1.06 0-2.08.14-3.05.4L7.3 4.75A12.6 12.6 0 0 1 12 4Zm0 4a4 4 0 0 1 4 4c0 .38-.05.74-.14 1.08l-1.63-1.63A2.5 2.5 0 0 0 10.55 9.3L8.92 7.67c.34-.09.7-.14 1.08-.14Z"
      />
    </svg>
  ) : (
    <svg width="14" height="14" viewBox="0 0 24 24" aria-hidden="true" style={{ display: "block" }}>
      <path
        fill="currentColor"
        d="M12 5c-5.5 0-9.9 3.6-11.5 8 1.6 4.4 6 8 11.5 8s9.9-3.6 11.5-8C21.9 8.6 17.5 5 12 5Zm0 13c-4.3 0-8-2.7-9.5-6 1.5-3.3 5.2-6 9.5-6s8 2.7 9.5 6c-1.5 3.3-5.2 6-9.5 6Zm0-10a4 4 0 1 0 0 8 4 4 0 0 0 0-8Zm0 6.5a2.5 2.5 0 1 1 0-5 2.5 2.5 0 0 1 0 5Z"
      />
    </svg>
  );
}

export default function App() {
  const [tab, setTab] = useState<Tab>("odoo");
  const [ctx, setCtx] = useState<OutlookMessageContext>({});
  const [meta, setMeta] = useState<OdooMeta | null>(null);
  const [links, setLinks] = useState<LinkEntry[]>([]);
  const [msg, setMsg] = useState<string | null>(null);

  // Thread oculto por defeito (só mostra ao clicar no ícone)
  const [showThread, setShowThread] = useState(false);

  // Ler contexto do email atual
  useEffect(() => {
    (async () => {
      try {
        const c = await getSelectedMessageContext();
        setCtx(c);
        if (!c.conversationId) {
          setMsg("Ainda não consegui ler o email via Office.js. (Troca de email e volta, ou reinicia o Outlook.)");
        } else {
          setMsg(null);
        }
      } catch (e: any) {
        setMsg(e?.message ?? String(e));
      }
    })();
  }, []);

  // Ler meta do Odoo
  useEffect(() => {
    (async () => {
      try {
        setMeta(await getOdooMeta());
      } catch (e: any) {
        setMsg((prev) => prev || (e?.message ?? String(e)));
      }
    })();
  }, []);

  async function refreshLinks() {
    if (!ctx.conversationId) return setLinks([]);
    try {
      setLinks(await getLinks(ctx.conversationId));
    } catch (e: any) {
      setMsg(e?.message ?? String(e));
    }
  }

  // Sempre que muda a conversationId (quando o Office atualizar o ctx), recarrega links
  useEffect(() => {
    refreshLinks();
    // eslint-disable-next-line react-hooks/exhaustive-deps
  }, [ctx.conversationId]);

  const fromLabel = useMemo(() => {
    if (ctx.fromName && ctx.fromEmail) return `${ctx.fromName} <${ctx.fromEmail}>`;
    return ctx.fromEmail || "—";
  }, [ctx.fromName, ctx.fromEmail]);

  return (
    <div style={S.shell}>
      <header style={S.header}>
        <div style={S.title}>Inbox CRM Cockpit</div>
        <div style={S.tabs}>
          <button style={tab === "odoo" ? S.tabA : S.tab} onClick={() => setTab("odoo")}>Odoo</button>
          <button style={tab === "ai" ? S.tabA : S.tab} onClick={() => setTab("ai")}>AI</button>
          <button style={tab === "settings" ? S.tabA : S.tab} onClick={() => setTab("settings")}>⚙️</button>
        </div>
      </header>

      <div style={S.card}>
        <div style={S.kv}><b>Assunto:</b> <span>{ctx.subject || "—"}</span></div>
        <div style={S.kv}><b>De:</b> <span>{fromLabel}</span></div>

        {/* Thread: oculto por defeito, só ícone pequeno para alternar */}
        <div style={{ ...S.kv, display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
          <div style={{ display: "flex", alignItems: "center", gap: 6 }}>
            <b>Thread:</b>
            {showThread ? (
              <span style={S.threadText}>{ctx.conversationId || "—"}</span>
            ) : (
              <span style={S.threadMuted}>oculto</span>
            )}
          </div>

          <button
            type="button"
            style={S.iconBtn}
            onClick={() => setShowThread((v) => !v)}
            title="Thread: usamos para agrupar itens relacionados nesta conversa"
            aria-label={showThread ? "Ocultar thread" : "Mostrar thread"}
          >
            <span style={{ display: "inline-flex", alignItems: "center", justifyContent: "center" }}>
              <EyeIcon off={showThread} />
            </span>
          </button>
        </div>
      </div>

      {tab === "odoo" && (
        <div style={S.card}>
          <div style={{ display: "flex", gap: 8 }}>
            <button
              style={S.btn}
              disabled={!ctx.conversationId}
              onClick={async () => {
                try {
                  await openCockpitDialog({
                    mode: "new",
                    conversationId: ctx.conversationId || "",
                    internetMessageId: ctx.internetMessageId || "",
                    subject: ctx.subject || "",
                    fromEmail: ctx.fromEmail || "",
                    fromName: ctx.fromName || "",
                    receivedAtIso: ctx.receivedDateTimeIso || "",
                  });
                  await refreshLinks();
                } catch (e: any) {
                  setMsg(e?.message ?? String(e));
                }
              }}
            >
              ➕ Criar
            </button>

            <button
              style={S.btn2}
              disabled={!ctx.conversationId}
              onClick={async () => {
                try {
                  await openCockpitDialog({
                    mode: "add",
                    conversationId: ctx.conversationId || "",
                    internetMessageId: ctx.internetMessageId || "",
                    subject: ctx.subject || "",
                    fromEmail: ctx.fromEmail || "",
                    fromName: ctx.fromName || "",
                    receivedAtIso: ctx.receivedDateTimeIso || "",
                  });
                  await refreshLinks();
                } catch (e: any) {
                  setMsg(e?.message ?? String(e));
                }
              }}
            >
              🔗 Ligar a existente
            </button>
          </div>

          <div style={{ marginTop: 12, borderTop: "1px solid #eee", paddingTop: 10 }}>
            <div style={{ fontWeight: 800 }}>Relacionados (esta conversa)</div>

            {links.length === 0 ? (
              <div style={{ marginTop: 8, color: "#666" }}>Nada ligado ainda.</div>
            ) : (
              <div style={{ marginTop: 8, display: "flex", flexDirection: "column", gap: 8 }}>
                {links.map((l, idx) => (
                  <RelatedRow
                    key={`${l.model}-${l.recordId}-${idx}`}
                    link={l}
                    meta={meta}
                    onEdit={async () => {
                      try {
                        await openCockpitDialog({
                          mode: "edit",
                          model: l.model,
                          recordId: String(l.recordId),
                          conversationId: ctx.conversationId || "",
                        });
                        await refreshLinks();
                      } catch (e: any) {
                        setMsg(e?.message ?? String(e));
                      }
                    }}
                  />
                ))}
              </div>
            )}

            <button style={S.linkBtn} onClick={refreshLinks}>↻ Recarregar</button>
          </div>

          {msg && <div style={S.alert}>{msg}</div>}
        </div>
      )}

      {tab === "ai" && (
        <div style={S.card}>
          <b>AI (MailMaestro)</b>
          <div style={{ marginTop: 8, color: "#666" }}>Próxima fase.</div>
        </div>
      )}

      {tab === "settings" && (
        <div style={S.card}>
          <b>Settings</b>
          <div style={{ marginTop: 8, color: "#666" }}>Vamos evoluir quando entrarem as keys e templates AI.</div>
        </div>
      )}

      <div style={{ marginTop: 10, fontSize: 12, color: "#666" }}>v4 • Dialog (janela) • Task form • Edit básico</div>
    </div>
  );
}

function RelatedRow({ link, meta, onEdit }: { link: LinkEntry; meta: OdooMeta | null; onEdit: () => void }) {
  const name = link.recordName || `${link.model} #${link.recordId}`;

  // Aceitar qualquer alias possível (baseUrl / webBaseUrl / url)
  const base =
    (meta as any)?.baseUrl ||
    (meta as any)?.webBaseUrl ||
    (meta as any)?.url ||
    "";

  const url = base
    ? `${String(base).replace(/\/+$/, "")}/web#id=${link.recordId}&model=${encodeURIComponent(link.model)}&view_type=form`
    : "";

  return (
    <div style={S.relRow}>
      <div style={{ flex: 1 }}>
        <div style={{ fontWeight: 800 }}>{name}</div>
        <div style={{ fontSize: 12, color: "#666" }}>{link.model} • #{link.recordId}</div>
      </div>
      {url ? (
        <a style={S.open} href={url} target="_blank" rel="noreferrer">Abrir</a>
      ) : (
        <span style={{ fontSize: 12, color: "#999" }}>—</span>
      )}
      <button style={S.smallBtn} onClick={onEdit}>Editar</button>
    </div>
  );
}

const S: Record<string, React.CSSProperties> = {
  shell: { fontFamily: "system-ui,Segoe UI,Arial", padding: 12 },
  header: { marginBottom: 10 },
  title: { fontWeight: 900, marginBottom: 8 },
  tabs: { display: "flex", gap: 8 },
  tab: { padding: "6px 10px", borderRadius: 8, border: "1px solid #ddd", background: "#fff", cursor: "pointer" },
  tabA: { padding: "6px 10px", borderRadius: 8, border: "1px solid #111", background: "#111", color: "#fff", cursor: "pointer" },

  card: { border: "1px solid #ddd", borderRadius: 12, padding: 12, background: "#fff", marginBottom: 10 },
  kv: { marginTop: 6 },

  threadText: { fontSize: 12, color: "#111", wordBreak: "break-all" },
  threadMuted: { fontSize: 12, color: "#666" },

  iconBtn: {
    width: 24,
    height: 24,
    padding: 0,
    borderRadius: 8,
    border: "1px solid #ddd",
    background: "#fff",
    cursor: "pointer",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    color: "#111",
  },

  btn: { flex: 1, padding: "10px 12px", borderRadius: 10, border: "1px solid #111", background: "#111", color: "#fff", cursor: "pointer" },
  btn2: { flex: 1, padding: "10px 12px", borderRadius: 10, border: "1px solid #ddd", background: "#f7f7f7", cursor: "pointer" },

  linkBtn: { marginTop: 10, padding: "6px 10px", borderRadius: 8, border: "1px solid #ddd", background: "#fff", cursor: "pointer" },
  alert: { marginTop: 10, padding: 10, borderRadius: 10, background: "#fff4f4", border: "1px solid #ffd1d1", whiteSpace: "pre-wrap" },

  relRow: { display: "flex", gap: 10, alignItems: "center", padding: 10, borderRadius: 10, border: "1px solid #eee", background: "#fafafa" },
  open: { textDecoration: "none", padding: "6px 10px", borderRadius: 8, border: "1px solid #ddd", background: "#fff", color: "#111" },
  smallBtn: { padding: "6px 10px", borderRadius: 8, border: "1px solid #111", background: "#111", color: "#fff", cursor: "pointer" },
};
