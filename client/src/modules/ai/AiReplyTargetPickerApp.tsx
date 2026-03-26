import React, { useEffect, useMemo, useState } from "react";
import { getRelatedEmailContext, type GroupTicketEntry, type LinkGroupEntry, type RelatedEmailEntry } from "@/api";
import { type AiReplyTargetSelection, requestCockpitHostAction } from "@/office";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import "../../global.css";

type PickerParams = {
  conversationId?: string;
  internetMessageId?: string;
  itemId?: string;
  subject?: string;
  fromEmail?: string;
  fromName?: string;
  receivedAtIso?: string;
  selectedEmailKey?: string;
};

function readParams(): PickerParams {
  const params = new URLSearchParams(window.location.search);
  return {
    conversationId: String(params.get("conversationId") || "").trim() || undefined,
    internetMessageId: String(params.get("internetMessageId") || "").trim() || undefined,
    itemId: String(params.get("itemId") || "").trim() || undefined,
    subject: String(params.get("subject") || "").trim() || undefined,
    fromEmail: String(params.get("fromEmail") || "").trim() || undefined,
    fromName: String(params.get("fromName") || "").trim() || undefined,
    receivedAtIso: String(params.get("receivedAtIso") || "").trim() || undefined,
    selectedEmailKey: String(params.get("selectedEmailKey") || "").trim() || undefined,
  };
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.emailKey || email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function normalizeMessageId(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function isCurrentEmail(email: RelatedEmailEntry, params: PickerParams): boolean {
  const currentItemId = String(params.itemId || "").trim();
  const currentMessageId = normalizeMessageId(params.internetMessageId);
  if (currentItemId && String(email.itemId || "").trim() === currentItemId) return true;
  if (currentMessageId && normalizeMessageId(email.internetMessageId) === currentMessageId) return true;
  return false;
}

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    year: "numeric",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function htmlToPlainText(html: string): string {
  return String(html || "")
    .replace(/<style[\s\S]*?<\/style>/gi, " ")
    .replace(/<script[\s\S]*?<\/script>/gi, " ")
    .replace(/<br\s*\/?>/gi, "\n")
    .replace(/<\/p>/gi, "\n")
    .replace(/<\/div>/gi, "\n")
    .replace(/<li[^>]*>/gi, "- ")
    .replace(/<\/li>/gi, "\n")
    .replace(/<[^>]+>/g, " ")
    .replace(/&nbsp;/gi, " ")
    .replace(/&amp;/gi, "&")
    .replace(/&lt;/gi, "<")
    .replace(/&gt;/gi, ">")
    .replace(/&#39;|&#039;/gi, "'")
    .replace(/&quot;/gi, "\"")
    .replace(/[ \t]{2,}/g, " ")
    .replace(/\n{3,}/g, "\n\n")
    .trim();
}

function escapeHtml(value: string): string {
  return String(value || "")
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function buildEmailPreviewHtml(email: RelatedEmailEntry | null): string {
  const html = String(email?.bodyHtml || "").trim();
  if (html) {
    return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      html, body { margin: 0; padding: 0; background: #ffffff; color: #172b4d; font: 14px/1.5 'Segoe UI', sans-serif; }
      body { padding: 18px; }
      img { max-width: 100%; height: auto; }
      table { max-width: 100%; }
      blockquote { margin-left: 0; padding-left: 12px; border-left: 3px solid #dbeafe; color: #475569; }
      pre { white-space: pre-wrap; word-break: break-word; }
    </style>
  </head>
  <body>${html}</body>
</html>`;
  }

  const text = String(email?.bodyText || "").trim();
  if (!text) return "";

  return `<!doctype html>
<html>
  <head>
    <meta charset="utf-8" />
    <style>
      html, body { margin: 0; padding: 0; background: #ffffff; color: #172b4d; font: 14px/1.55 'Segoe UI', sans-serif; }
      body { padding: 18px; }
      pre { margin: 0; white-space: pre-wrap; word-break: break-word; font: inherit; }
    </style>
  </head>
  <body><pre>${escapeHtml(text)}</pre></body>
</html>`;
}

function buildSnippet(email: RelatedEmailEntry): string {
  const source = String(email.bodyText || "").trim() || htmlToPlainText(String(email.bodyHtml || ""));
  return source.length > 180 ? `${source.slice(0, 177).trim()}...` : source;
}

function closeWindow() {
  void requestCockpitHostAction({ type: "close" });
}

function sendResult(selection: AiReplyTargetSelection) {
  try {
    const OfficeAny = (window as any).Office;
    if (typeof OfficeAny?.context?.ui?.messageParent === "function") {
      OfficeAny.context.ui.messageParent(JSON.stringify({ type: "dialog-result", result: selection }));
      return;
    }
  } catch {
    // fall through
  }
  window.close();
}

export default function AiReplyTargetPickerApp() {
  const params = useMemo(() => readParams(), []);
  const [emails, setEmails] = useState<RelatedEmailEntry[]>([]);
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [tickets, setTickets] = useState<GroupTicketEntry[]>([]);
  const [selectedEmailKey, setSelectedEmailKey] = useState(params.selectedEmailKey || "");
  const [search, setSearch] = useState("");
  const [loading, setLoading] = useState(true);
  const [error, setError] = useState("");

  useEffect(() => {
    void (async () => {
      try {
        const settings = await getSettings();
        applySkin(settings.skin || "soft");
      } catch {
        applySkin("soft");
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    void (async () => {
      setLoading(true);
      setError("");
      try {
        const related = await getRelatedEmailContext({
          conversationId: params.conversationId,
          internetMessageId: params.internetMessageId,
          itemId: params.itemId,
          subject: params.subject,
          fromEmail: params.fromEmail,
          fromName: params.fromName,
          receivedAtIso: params.receivedAtIso,
        });
        if (cancelled) return;
        const nextEmails = Array.isArray(related.emails) ? related.emails : [];
        setEmails(nextEmails);
        setGroups(Array.isArray(related.groups) ? related.groups : []);
        setTickets(Array.isArray(related.tickets) ? related.tickets : []);
        setSelectedEmailKey((current) => {
          if (current && nextEmails.some((entry) => makeEmailKey(entry) === current)) return current;
          if (params.selectedEmailKey && nextEmails.some((entry) => makeEmailKey(entry) === params.selectedEmailKey)) return params.selectedEmailKey;
          const firstExternal = nextEmails.find((entry) => !isCurrentEmail(entry, params));
          return makeEmailKey(firstExternal || nextEmails[0] || {});
        });
      } catch (fetchError: any) {
        if (!cancelled) {
          setError(String(fetchError?.message || fetchError || "Falha a carregar emails relacionados."));
        }
      } finally {
        if (!cancelled) setLoading(false);
      }
    })();

    return () => {
      cancelled = true;
    };
  }, [params]);

  const filteredEmails = useMemo(() => {
    const q = String(search || "").trim().toLowerCase();
    const rows = [...emails].sort((a, b) =>
      String(b.messageDateIso || b.receivedAtIso || "").localeCompare(String(a.messageDateIso || a.receivedAtIso || ""))
    );
    if (!q) return rows;
    return rows.filter((email) => {
      const haystack = [
        email.subject,
        email.fromName,
        email.fromEmail,
        buildSnippet(email),
      ].join(" ").toLowerCase();
      return haystack.includes(q);
    });
  }, [emails, search]);

  const selectedEmail = useMemo(
    () => filteredEmails.find((entry) => makeEmailKey(entry) === selectedEmailKey) || emails.find((entry) => makeEmailKey(entry) === selectedEmailKey) || filteredEmails[0] || emails[0] || null,
    [emails, filteredEmails, selectedEmailKey]
  );

  const selectedPreviewHtml = useMemo(() => buildEmailPreviewHtml(selectedEmail), [selectedEmail]);

  function handleUseSelectedEmail() {
    if (!selectedEmail) return;
    sendResult({
      emailKey: makeEmailKey(selectedEmail),
      itemId: selectedEmail.itemId,
      emailWebLink: selectedEmail.emailWebLink,
      internetMessageId: selectedEmail.internetMessageId,
      conversationId: selectedEmail.conversationId,
      subject: selectedEmail.subject,
      fromEmail: selectedEmail.fromEmail,
      fromName: selectedEmail.fromName,
      messageDateIso: selectedEmail.messageDateIso,
      receivedAtIso: selectedEmail.receivedAtIso,
      bodyText: selectedEmail.bodyText,
      bodyHtml: selectedEmail.bodyHtml,
    });
  }

  return (
    <div style={styles.root}>
      <div style={styles.header}>
        <div style={styles.headerText}>
          <div style={styles.kicker}>IA</div>
          <div style={styles.title}>Escolher Email-Alvo</div>
          <div style={styles.subtitle}>Seleciona o email guardado ao qual queres responder usando o contexto do caso atual.</div>
        </div>
        <div style={styles.headerActions}>
          <button type="button" style={styles.secondaryBtn} onClick={closeWindow}>Fechar</button>
          <button type="button" style={styles.primaryBtn} disabled={!selectedEmail} onClick={handleUseSelectedEmail}>Usar este email</button>
        </div>
      </div>

      <div style={styles.contextRow}>
        <div style={styles.contextBlock}>
          <div style={styles.contextLabel}>Email atual</div>
          <div style={styles.contextValue}>{params.subject || "(sem assunto)"}</div>
        </div>
        <div style={styles.contextTags}>
          {tickets.slice(0, 4).map((ticket) => (
            <span key={ticket.id} style={styles.ticketTag}>{ticket.code}</span>
          ))}
          {groups.slice(0, 4).map((group) => (
            <span key={group.id} style={styles.groupTag}>{group.name}</span>
          ))}
        </div>
      </div>

      {error ? <div style={styles.errorBox}>{error}</div> : null}

      <div style={styles.shell}>
        <section style={styles.listPanel}>
          <div style={styles.panelHeader}>
            <div style={styles.panelTitle}>Emails relacionados</div>
            <div style={styles.panelMeta}>{filteredEmails.length}</div>
          </div>
          <input
            style={styles.searchInput}
            value={search}
            onChange={(event) => setSearch(event.target.value)}
            placeholder="Pesquisar por assunto, remetente ou texto..."
          />
          <div style={styles.listBody}>
            {loading ? <PanelState compact tone="loading" title="A carregar emails" description="A listar os emails relacionados com o ticket/grupo atual." /> : null}
            {!loading && !filteredEmails.length ? (
              <PanelState compact tone="info" title="Sem emails relacionados" description="Ainda nao ha emails guardados suficientes para selecionar um alvo." />
            ) : null}
            {!loading && filteredEmails.map((email) => {
              const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
              const current = isCurrentEmail(email, params);
              return (
                <button
                  key={makeEmailKey(email)}
                  type="button"
                  style={active ? styles.emailCardActive : styles.emailCard}
                  onClick={() => setSelectedEmailKey(makeEmailKey(email))}
                >
                  <div style={styles.emailCardTop}>
                    <div style={styles.emailSubject}>{email.subject || "(sem assunto)"}</div>
                    {current ? <span style={styles.currentTag}>Atual</span> : null}
                  </div>
                  <div style={styles.emailMeta}>{email.fromName || email.fromEmail || "--"} · {formatDate(email.messageDateIso || email.receivedAtIso) || "--"}</div>
                  <div style={styles.emailSnippet}>{buildSnippet(email) || "Sem preview curto disponivel."}</div>
                </button>
              );
            })}
          </div>
        </section>

        <section style={styles.previewPanel}>
          <div style={styles.panelHeader}>
            <div style={styles.previewHeaderMain}>
              <div style={styles.panelTitle}>Preview</div>
              {selectedEmail ? (
                <div style={styles.previewMeta}>
                  <span>{selectedEmail.fromName || selectedEmail.fromEmail || "--"}</span>
                  <span>{formatDate(selectedEmail.messageDateIso || selectedEmail.receivedAtIso) || "--"}</span>
                </div>
              ) : null}
            </div>
            {selectedEmail?.itemId || selectedEmail?.emailWebLink ? (
              <button
                type="button"
                style={styles.secondaryBtn}
                onClick={() => void requestCockpitHostAction({ type: "open-email", itemId: selectedEmail.itemId, emailWebLink: selectedEmail.emailWebLink })}
              >
                Abrir no Outlook
              </button>
            ) : null}
          </div>

          {!selectedEmail ? (
            <PanelState compact tone="info" title="Sem email selecionado" description="Escolhe um email na coluna da esquerda para abrir o preview." />
          ) : !selectedPreviewHtml ? (
            <PanelState compact tone="info" title="Preview indisponivel" description="Este email ainda nao tem corpo guardado para preview detalhado." />
          ) : (
            <>
              <div style={styles.previewSummary}>
                <div style={styles.previewSummarySubject}>{selectedEmail.subject || "(sem assunto)"}</div>
                <div style={styles.previewSummaryLine}>De: {selectedEmail.fromName || "--"}{selectedEmail.fromEmail ? ` <${selectedEmail.fromEmail}>` : ""}</div>
              </div>
              <div style={styles.previewFrame}>
                <iframe title={selectedEmail.subject || "Preview do email"} srcDoc={selectedPreviewHtml} style={styles.previewIframe} sandbox="" />
              </div>
            </>
          )}
        </section>
      </div>
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: {
    height: "100vh",
    boxSizing: "border-box",
    padding: "20px",
    display: "grid",
    gridTemplateRows: "auto auto minmax(0, 1fr)",
    gap: "14px",
    background: "linear-gradient(180deg, rgba(248,250,252,0.98) 0%, rgba(239,244,252,0.94) 100%)",
    color: "#0f172a",
    fontFamily: "var(--iccc-font, 'Segoe UI', sans-serif)",
  },
  header: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: "16px",
    padding: "12px 16px",
    borderRadius: "18px",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.94)",
    boxShadow: "0 12px 28px rgba(15, 23, 42, 0.06)",
  },
  headerText: {
    display: "grid",
    gap: "4px",
    minWidth: 0,
  },
  kicker: {
    fontSize: "10px",
    fontWeight: 700,
    letterSpacing: "0.08em",
    textTransform: "uppercase",
    color: "#64748b",
  },
  title: {
    fontSize: "24px",
    fontWeight: 800,
    color: "#0f172a",
  },
  subtitle: {
    fontSize: "13px",
    color: "#475569",
    maxWidth: "780px",
  },
  headerActions: {
    display: "flex",
    alignItems: "center",
    gap: "8px",
  },
  primaryBtn: {
    borderRadius: "999px",
    border: "1px solid rgba(37, 99, 235, 0.28)",
    background: "linear-gradient(180deg, rgba(59,130,246,0.96) 0%, rgba(29,78,216,0.9) 100%)",
    color: "#ffffff",
    fontSize: "12px",
    fontWeight: 800,
    padding: "10px 16px",
    cursor: "pointer",
    boxShadow: "0 8px 18px rgba(37, 99, 235, 0.22)",
  },
  secondaryBtn: {
    borderRadius: "999px",
    border: "1px solid rgba(15, 23, 42, 0.12)",
    background: "#ffffff",
    color: "#0f172a",
    fontSize: "12px",
    fontWeight: 700,
    padding: "9px 14px",
    cursor: "pointer",
  },
  contextRow: {
    display: "flex",
    justifyContent: "space-between",
    alignItems: "center",
    gap: "12px",
    padding: "10px 14px",
    borderRadius: "16px",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.9)",
  },
  contextBlock: {
    minWidth: 0,
  },
  contextLabel: {
    fontSize: "10px",
    fontWeight: 700,
    textTransform: "uppercase",
    letterSpacing: "0.08em",
    color: "#64748b",
    marginBottom: "2px",
  },
  contextValue: {
    fontSize: "14px",
    fontWeight: 700,
    color: "#0f172a",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  contextTags: {
    display: "flex",
    flexWrap: "wrap",
    gap: "6px",
    justifyContent: "flex-end",
  },
  ticketTag: {
    borderRadius: "999px",
    padding: "4px 8px",
    fontSize: "10px",
    fontWeight: 800,
    background: "rgba(37, 99, 235, 0.1)",
    color: "#1d4ed8",
    border: "1px solid rgba(37, 99, 235, 0.18)",
  },
  groupTag: {
    borderRadius: "999px",
    padding: "4px 8px",
    fontSize: "10px",
    fontWeight: 700,
    background: "rgba(15, 23, 42, 0.05)",
    color: "#334155",
    border: "1px solid rgba(15, 23, 42, 0.08)",
  },
  errorBox: {
    padding: "10px 14px",
    borderRadius: "14px",
    border: "1px solid rgba(220, 38, 38, 0.18)",
    background: "rgba(254, 242, 242, 0.95)",
    color: "#991b1b",
    fontSize: "13px",
    fontWeight: 600,
  },
  shell: {
    minHeight: 0,
    display: "grid",
    gridTemplateColumns: "320px minmax(0, 1fr)",
    gap: "14px",
  },
  listPanel: {
    minHeight: 0,
    display: "grid",
    gridTemplateRows: "auto auto minmax(0, 1fr)",
    gap: "10px",
    padding: "14px",
    borderRadius: "18px",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.94)",
    boxShadow: "0 12px 28px rgba(15, 23, 42, 0.06)",
  },
  previewPanel: {
    minHeight: 0,
    display: "grid",
    gridTemplateRows: "auto auto minmax(0, 1fr)",
    gap: "10px",
    padding: "14px",
    borderRadius: "18px",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.94)",
    boxShadow: "0 12px 28px rgba(15, 23, 42, 0.06)",
  },
  panelHeader: {
    display: "flex",
    alignItems: "center",
    justifyContent: "space-between",
    gap: "10px",
  },
  previewHeaderMain: {
    display: "grid",
    gap: "2px",
    minWidth: 0,
  },
  panelTitle: {
    fontSize: "16px",
    fontWeight: 800,
    color: "#0f172a",
  },
  panelMeta: {
    minWidth: "28px",
    height: "28px",
    borderRadius: "999px",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    background: "rgba(37, 99, 235, 0.08)",
    color: "#1d4ed8",
    fontSize: "11px",
    fontWeight: 800,
  },
  searchInput: {
    width: "100%",
    boxSizing: "border-box",
    borderRadius: "12px",
    border: "1px solid rgba(148, 163, 184, 0.28)",
    background: "#ffffff",
    color: "#0f172a",
    fontSize: "13px",
    padding: "10px 12px",
    outline: "none",
  },
  listBody: {
    minHeight: 0,
    overflowY: "auto",
    display: "grid",
    gap: "8px",
    paddingRight: "4px",
  },
  emailCard: {
    textAlign: "left",
    borderRadius: "14px",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "#ffffff",
    padding: "12px",
    display: "grid",
    gap: "6px",
    cursor: "pointer",
  },
  emailCardActive: {
    textAlign: "left",
    borderRadius: "14px",
    border: "1px solid rgba(37, 99, 235, 0.24)",
    background: "rgba(219, 234, 254, 0.72)",
    padding: "12px",
    display: "grid",
    gap: "6px",
    cursor: "pointer",
    boxShadow: "inset 0 0 0 1px rgba(37, 99, 235, 0.08)",
  },
  emailCardTop: {
    display: "flex",
    alignItems: "flex-start",
    justifyContent: "space-between",
    gap: "8px",
  },
  emailSubject: {
    fontSize: "13px",
    fontWeight: 800,
    lineHeight: 1.35,
    color: "#0f172a",
  },
  currentTag: {
    borderRadius: "999px",
    padding: "2px 7px",
    background: "rgba(15, 23, 42, 0.06)",
    color: "#475569",
    fontSize: "9px",
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    flexShrink: 0,
  },
  emailMeta: {
    fontSize: "11px",
    color: "#475569",
  },
  emailSnippet: {
    fontSize: "11px",
    lineHeight: 1.45,
    color: "#334155",
    display: "-webkit-box",
    WebkitLineClamp: 4,
    WebkitBoxOrient: "vertical",
    overflow: "hidden",
  },
  previewMeta: {
    display: "flex",
    flexWrap: "wrap",
    gap: "10px",
    fontSize: "11px",
    color: "#64748b",
  },
  previewSummary: {
    display: "grid",
    gap: "4px",
    padding: "10px 12px",
    borderRadius: "14px",
    background: "rgba(248,250,252,0.92)",
    border: "1px solid rgba(15, 23, 42, 0.06)",
  },
  previewSummarySubject: {
    fontSize: "16px",
    fontWeight: 800,
    color: "#0f172a",
  },
  previewSummaryLine: {
    fontSize: "12px",
    color: "#475569",
  },
  previewFrame: {
    minHeight: 0,
    height: "100%",
    borderRadius: "16px",
    overflow: "hidden",
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "#f8fafc",
  },
  previewIframe: {
    width: "100%",
    height: "100%",
    border: "none",
    display: "block",
    background: "#ffffff",
  },
};
