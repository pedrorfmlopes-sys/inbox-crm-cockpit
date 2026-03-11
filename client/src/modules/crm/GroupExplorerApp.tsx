import React, { useEffect, useMemo, useRef, useState } from "react";
import {
  deleteGroupDocument,
  getGroupDocuments,
  getGroupEmails,
  listLinkGroups,
  removeEmailFromLinkGroup,
  type GroupDocumentEntry,
  type LinkGroupEntry,
  type RelatedEmailEntry,
} from "@/api";
import { addBase64AttachmentToCompose, openLinkedOutlookEmail } from "@/office";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import "../../global.css";

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", {
    day: "2-digit",
    month: "2-digit",
    hour: "2-digit",
    minute: "2-digit",
  });
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function makeDocumentKey(document: Partial<GroupDocumentEntry>): string {
  return String(document?.id || document?.storagePathHint || document?.name || "");
}

function formatBytes(value: number | undefined): string {
  const size = Number(value || 0);
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function closeExplorer() {
  try {
    (window as any).Office?.context?.ui?.messageParent?.("close");
  } catch {
    // ignore
  }
  try {
    window.close();
  } catch {
    // ignore
  }
}

function normalizeProvider(value: string | undefined): "cloud" | "local" | "onedrive" {
  const normalized = String(value || "").trim().toLowerCase();
  if (normalized === "local" || normalized === "onedrive") return normalized;
  return "cloud";
}

function providerLabel(value: string | undefined): string {
  const provider = normalizeProvider(value);
  if (provider === "onedrive") return "OneDrive / SharePoint";
  if (provider === "local") return "Pasta local do utilizador";
  return "Cockpit Cloud";
}

function readExplorerParams() {
  const params = new URLSearchParams(window.location.search);
  return {
    groupId: String(params.get("groupId") || "").trim(),
    emailKey: String(params.get("emailKey") || "").trim(),
    documentId: String(params.get("documentId") || "").trim(),
  };
}

export default function GroupExplorerApp(): JSX.Element {
  const initial = useMemo(() => readExplorerParams(), []);
  const downloadAnchorRef = useRef<HTMLAnchorElement | null>(null);
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [selectedGroupId, setSelectedGroupId] = useState(initial.groupId);
  const [selectedEmailKey, setSelectedEmailKey] = useState(initial.emailKey);
  const [selectedDocumentId, setSelectedDocumentId] = useState(initial.documentId);
  const [groupEmails, setGroupEmails] = useState<RelatedEmailEntry[]>([]);
  const [groupDocuments, setGroupDocuments] = useState<GroupDocumentEntry[]>([]);
  const [loadingGroups, setLoadingGroups] = useState(true);
  const [loadingEmails, setLoadingEmails] = useState(false);
  const [loadingDocuments, setLoadingDocuments] = useState(false);
  const [error, setError] = useState<string | null>(null);
  const [notice, setNotice] = useState<string | null>(null);
  const [busy, setBusy] = useState(false);

  useEffect(() => {
    (async () => {
      try {
        const settings = await getSettings();
        if (settings.skinId) applySkin(settings.skinId);
      } catch {
        // ignore
      }
    })();
  }, []);

  useEffect(() => {
    let cancelled = false;
    setLoadingGroups(true);
    setError(null);
    listLinkGroups("/")
      .then((nextGroups) => {
        if (cancelled) return;
        setGroups(nextGroups);
        setSelectedGroupId((current) => {
          if (current && nextGroups.some((group) => group.id === current)) return current;
          if (initial.groupId && nextGroups.some((group) => group.id === initial.groupId)) return initial.groupId;
          return nextGroups[0]?.id || "";
        });
      })
      .catch((nextError: any) => {
        if (cancelled) return;
        setError(nextError?.message || "Nao foi possivel carregar os grupos.");
      })
      .finally(() => {
        if (!cancelled) setLoadingGroups(false);
      });

    return () => {
      cancelled = true;
    };
  }, [initial.groupId]);

  useEffect(() => {
    if (!selectedGroupId) {
      setGroupEmails([]);
      setGroupDocuments([]);
      return;
    }

    let cancelled = false;
    setLoadingEmails(true);
    setLoadingDocuments(true);
    setError(null);

    Promise.all([getGroupEmails(selectedGroupId), getGroupDocuments(selectedGroupId)])
      .then(([emails, documents]) => {
        if (cancelled) return;
        setGroupEmails(emails);
        setGroupDocuments(documents);
        setSelectedEmailKey((current) => {
          if (current && emails.some((email) => makeEmailKey(email) === current)) return current;
          if (initial.emailKey && emails.some((email) => makeEmailKey(email) === initial.emailKey)) return initial.emailKey;
          return emails[0] ? makeEmailKey(emails[0]) : "";
        });
        setSelectedDocumentId((current) => {
          if (current && documents.some((document) => makeDocumentKey(document) === current)) return current;
          if (initial.documentId && documents.some((document) => makeDocumentKey(document) === initial.documentId)) return initial.documentId;
          return documents[0] ? makeDocumentKey(documents[0]) : "";
        });
      })
      .catch((nextError: any) => {
        if (cancelled) return;
        setError(nextError?.message || "Nao foi possivel carregar os dados do grupo.");
      })
      .finally(() => {
        if (cancelled) return;
        setLoadingEmails(false);
        setLoadingDocuments(false);
      });

    return () => {
      cancelled = true;
    };
  }, [initial.documentId, initial.emailKey, selectedGroupId]);

  const selectedGroup = useMemo(
    () => groups.find((group) => group.id === selectedGroupId) || null,
    [groups, selectedGroupId]
  );
  const selectedEmail = useMemo(
    () => groupEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || groupEmails[0] || null,
    [groupEmails, selectedEmailKey]
  );
  const selectedDocument = useMemo(
    () => groupDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || groupDocuments[0] || null,
    [groupDocuments, selectedDocumentId]
  );
  const selectedDocumentPreview = useMemo(() => {
    if (!selectedDocument?.contentBase64) return null;
    const contentType = String(selectedDocument.contentType || "").toLowerCase();
    const dataUrl = `data:${selectedDocument.contentType || "application/octet-stream"};base64,${selectedDocument.contentBase64}`;
    if (contentType.startsWith("image/")) return { kind: "image" as const, dataUrl };
    if (contentType === "application/pdf") return { kind: "pdf" as const, dataUrl };
    if (contentType.startsWith("text/") || contentType.includes("json") || contentType.includes("xml")) {
      try {
        return { kind: "text" as const, text: globalThis.atob(selectedDocument.contentBase64) };
      } catch {
        return { kind: "unsupported" as const };
      }
    }
    return { kind: "unsupported" as const };
  }, [selectedDocument]);

  async function refreshCurrentGroup() {
    if (!selectedGroupId) return;
    setLoadingEmails(true);
    setLoadingDocuments(true);
    setError(null);
    try {
      const [emails, documents] = await Promise.all([getGroupEmails(selectedGroupId), getGroupDocuments(selectedGroupId)]);
      setGroupEmails(emails);
      setGroupDocuments(documents);
      setNotice("Explorador atualizado.");
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel atualizar o explorador.");
    } finally {
      setLoadingEmails(false);
      setLoadingDocuments(false);
    }
  }

  async function handleOpenEmail(email: RelatedEmailEntry) {
    const opened = await openLinkedOutlookEmail({
      itemId: email.itemId,
      emailWebLink: email.emailWebLink,
    });
    if (!opened) setNotice("Este email ainda nao tem abertura direta disponivel.");
  }

  async function handleRemoveEmail(email: RelatedEmailEntry) {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      const persistentEmailKey = String(email.id || "").startsWith("email_") ? undefined : String(email.id || "").trim() || undefined;
      await removeEmailFromLinkGroup(selectedGroup.id, {
        emailKey: persistentEmailKey,
        itemId: email.itemId,
        internetMessageId: email.internetMessageId,
        conversationId: email.conversationId,
        subject: email.subject,
        fromEmail: email.fromEmail,
        receivedAtIso: email.receivedAtIso || email.messageDateIso,
      });
      await refreshCurrentGroup();
      setNotice("Email removido do grupo.");
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel remover o email do grupo.");
    } finally {
      setBusy(false);
    }
  }

  function handleDownloadDocument(document: GroupDocumentEntry) {
    const base64 = String(document.contentBase64 || "").trim();
    if (!base64) {
      setNotice("Este documento nao tem conteudo disponivel para download.");
      return;
    }
    const bytes = globalThis.atob(base64);
    const buffer = new Array(bytes.length);
    for (let index = 0; index < bytes.length; index += 1) buffer[index] = bytes.charCodeAt(index);
    const blob = new Blob([new Uint8Array(buffer)], { type: document.contentType || "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const anchor = downloadAnchorRef.current || globalThis.document.createElement("a");
    downloadAnchorRef.current = anchor;
    anchor.href = url;
    anchor.download = document.name || "documento";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 2000);
  }

  async function handleAttachDocument(document: GroupDocumentEntry) {
    try {
      await addBase64AttachmentToCompose(document.name || "documento", String(document.contentBase64 || ""));
      setNotice(`Documento "${document.name}" anexado ao email em edicao.`);
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel anexar o documento.");
    }
  }

  async function handleDeleteDocument(document: GroupDocumentEntry) {
    if (!selectedGroup) return;
    setBusy(true);
    try {
      await deleteGroupDocument(selectedGroup.id, document.id);
      await refreshCurrentGroup();
      setNotice(`Documento "${document.name}" removido.`);
    } catch (nextError: any) {
      setError(nextError?.message || "Nao foi possivel remover o documento.");
    } finally {
      setBusy(false);
    }
  }

  return (
    <div style={styles.root}>
      <header style={styles.header}>
        <div style={styles.headerCopy}>
          <div style={styles.eyebrow}>Explorador documental</div>
          <div style={styles.title}>Grupos</div>
          <div style={styles.subtitle}>
            Navega pelos emails e documentos guardados no grupo selecionado, com mais espaÃ§o do que no taskpane.
          </div>
        </div>
        <div style={styles.headerActions}>
          <div style={styles.selectWrap}>
            <select
              style={styles.select}
              value={selectedGroupId}
              onChange={(event) => {
                setSelectedGroupId(event.target.value);
                setSelectedEmailKey("");
                setSelectedDocumentId("");
                setNotice(null);
              }}
            >
              {groups.map((group) => (
                <option key={group.id} value={group.id}>
                  {group.name}
                </option>
              ))}
            </select>
          </div>
          <button type="button" style={styles.iconBtn} onClick={() => void refreshCurrentGroup()} disabled={loadingGroups || loadingEmails || loadingDocuments}>
            <Icons.RefreshCw size={14} />
          </button>
          <button type="button" style={styles.closeBtn} onClick={closeExplorer}>
            Fechar
          </button>
        </div>
      </header>

      {error ? <PanelState compact tone="error" title="Falha no explorador" description={error} /> : null}
      {notice ? <PanelState compact tone="info" title="Explorador" description={notice} /> : null}

      {loadingGroups ? <PanelState compact tone="loading" title="A carregar grupos" description="Estamos a preparar o explorador documental." /> : null}
      {!loadingGroups && !selectedGroup ? <PanelState compact tone="info" title="Sem grupos" description="Ainda nao existem grupos manuais disponiveis para este explorador." /> : null}

      {selectedGroup ? (
        <>
          <section style={styles.summaryCard}>
            <div style={styles.metricGrid}>
              <div style={styles.metric}>
                <span style={styles.metricLabel}>Grupo</span>
                <span style={styles.metricValue}>{selectedGroup.name}</span>
              </div>
              <div style={styles.metric}>
                <span style={styles.metricLabel}>Provider</span>
                <span style={styles.metricValue}>{providerLabel(groupDocuments[0]?.storageProvider)}</span>
              </div>
              <div style={styles.metric}>
                <span style={styles.metricLabel}>Emails</span>
                <span style={styles.metricValue}>{groupEmails.length}</span>
              </div>
              <div style={styles.metric}>
                <span style={styles.metricLabel}>Documentos</span>
                <span style={styles.metricValue}>{groupDocuments.length}</span>
              </div>
            </div>
          </section>

          <section style={styles.section}>
            <div style={styles.sectionHeader}>
              <div>
                <div style={styles.sectionTitle}>Emails</div>
                <div style={styles.sectionHint}>Emails ligados a este grupo, com acesso rÃ¡pido para abrir e limpar memberships.</div>
              </div>
            </div>
            <div style={styles.scrollBlock}>
              {loadingEmails && !groupEmails.length ? <PanelState compact tone="loading" title="A carregar emails" description="A listar os emails do grupo." /> : null}
              {!loadingEmails && !groupEmails.length ? <PanelState compact tone="info" title="Sem emails" description="Este grupo ainda nao tem emails visiveis no explorador." /> : null}
              {groupEmails.map((email) => {
                const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                const canOpen = Boolean(email.itemId || email.emailWebLink);
                return (
                  <div key={makeEmailKey(email)} style={active ? styles.rowActive : styles.row}>
                    <button type="button" style={styles.rowMain} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                      <div style={styles.rowTitle}>{email.subject || "(sem assunto)"}</div>
                      <div style={styles.rowMeta}>
                        <span>{email.fromName || email.fromEmail || "(sem remetente)"}</span>
                        {formatDate(email.messageDateIso || email.receivedAtIso) ? <span>{formatDate(email.messageDateIso || email.receivedAtIso)}</span> : null}
                      </div>
                    </button>
                    <div style={styles.rowActions}>
                      <button type="button" style={styles.iconBtn} onClick={() => void handleOpenEmail(email)} disabled={!canOpen}>
                        <Icons.MessageSquare size={12} />
                      </button>
                      <button type="button" style={styles.iconBtnDanger} onClick={() => void handleRemoveEmail(email)} disabled={busy}>
                        <Icons.Trash size={12} />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </section>

          <section style={styles.section}>
            <div style={styles.sectionHeader}>
              <div>
                <div style={styles.sectionTitle}>Documentos</div>
                <div style={styles.sectionHint}>Vault central do grupo. Daqui podes visualizar, descarregar, anexar e apagar.</div>
              </div>
            </div>
            <div style={styles.scrollBlock}>
              {loadingDocuments && !groupDocuments.length ? <PanelState compact tone="loading" title="A carregar documentos" description="A listar os documentos guardados." /> : null}
              {!loadingDocuments && !groupDocuments.length ? <PanelState compact tone="info" title="Sem documentos" description="Este grupo ainda nao tem documentos guardados no Cockpit Cloud." /> : null}
              {groupDocuments.map((document) => {
                const active = makeDocumentKey(document) === makeDocumentKey(selectedDocument || {});
                return (
                  <div key={makeDocumentKey(document)} style={active ? styles.rowActive : styles.row}>
                    <button type="button" style={styles.rowMain} onClick={() => setSelectedDocumentId(makeDocumentKey(document))}>
                      <div style={styles.rowTitle}>{document.name}</div>
                      <div style={styles.rowMeta}>
                        <span>{document.contentType || "Documento"}</span>
                        {formatBytes(document.size) ? <span>{formatBytes(document.size)}</span> : null}
                        {document.sourceEmailSubject ? <span>{document.sourceEmailSubject}</span> : null}
                      </div>
                    </button>
                    <div style={styles.rowActions}>
                      <button type="button" style={styles.iconBtn} onClick={() => handleDownloadDocument(document)} disabled={!document.contentBase64}>
                        <Icons.Download size={12} />
                      </button>
                      <button type="button" style={styles.iconBtn} onClick={() => void handleAttachDocument(document)} disabled={!document.contentBase64}>
                        <Icons.Upload size={12} />
                      </button>
                      <button type="button" style={styles.iconBtnDanger} onClick={() => void handleDeleteDocument(document)} disabled={busy}>
                        <Icons.Trash size={12} />
                      </button>
                    </div>
                  </div>
                );
              })}
            </div>
          </section>

          <section style={styles.section}>
            <div style={styles.sectionHeader}>
              <div>
                <div style={styles.sectionTitle}>Preview</div>
                <div style={styles.sectionHint}>Vista rÃ¡pida do documento selecionado.</div>
              </div>
            </div>
            {!selectedDocument ? <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe um documento acima para o visualizar." /> : null}
            {selectedDocument && selectedDocumentPreview?.kind === "image" ? (
              <div style={styles.previewFrame}>
                <img src={selectedDocumentPreview.dataUrl} alt={selectedDocument.name} style={styles.previewImage} />
              </div>
            ) : null}
            {selectedDocument && selectedDocumentPreview?.kind === "pdf" ? (
              <div style={styles.previewFrame}>
                <iframe title={selectedDocument.name} src={selectedDocumentPreview.dataUrl} style={styles.previewIframe} />
              </div>
            ) : null}
            {selectedDocument && selectedDocumentPreview?.kind === "text" ? (
              <pre style={styles.previewText}>{selectedDocumentPreview.text}</pre>
            ) : null}
            {selectedDocument && (!selectedDocumentPreview || selectedDocumentPreview.kind === "unsupported") ? (
              <PanelState compact tone="info" title="Preview nÃ£o disponÃ­vel" description="Este documento pode ser descarregado ou anexado, mas ainda nÃ£o tem preview interno." />
            ) : null}
          </section>
        </>
      ) : null}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: {
    minHeight: "100vh",
    background: "var(--iccc-bg, #edf2f7)",
    color: "var(--iccc-text, #172b4d)",
    fontFamily: "var(--iccc-font, 'Segoe UI', sans-serif)",
    padding: 14,
    display: "grid",
    gap: 12,
  },
  header: {
    display: "grid",
    gap: 10,
    padding: 12,
    borderRadius: 14,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.86)",
    boxShadow: "0 10px 28px rgba(15, 23, 42, 0.08)",
  },
  headerCopy: {
    display: "grid",
    gap: 4,
  },
  eyebrow: {
    fontSize: 11,
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.08em",
    color: "#5b6b83",
  },
  title: {
    fontSize: 24,
    fontWeight: 900,
    color: "#0f172a",
  },
  subtitle: {
    fontSize: 12,
    color: "#607086",
    lineHeight: 1.45,
  },
  headerActions: {
    display: "flex",
    gap: 8,
    alignItems: "center",
    flexWrap: "wrap",
  },
  selectWrap: {
    flex: "1 1 240px",
    minWidth: 180,
    border: "none",
    padding: 0,
    background: "transparent",
  },
  select: {
    width: "100%",
    borderRadius: 999,
    border: "1px solid rgba(15, 23, 42, 0.12)",
    background: "rgba(248,250,252,0.95)",
    color: "#172b4d",
    padding: "9px 14px",
    fontSize: 12,
    fontWeight: 700,
    outline: "none",
  },
  closeBtn: {
    borderRadius: 999,
    border: "none",
    background: "#1d4ed8",
    color: "#fff",
    padding: "9px 16px",
    fontSize: 12,
    fontWeight: 800,
    cursor: "pointer",
  },
  summaryCard: {
    padding: 12,
    borderRadius: 14,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.86)",
    boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)",
  },
  metricGrid: {
    display: "grid",
    gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
    gap: 8,
  },
  metric: {
    display: "grid",
    gap: 2,
    padding: 10,
    borderRadius: 12,
    background: "rgba(15, 23, 42, 0.03)",
  },
  metricLabel: {
    fontSize: 10,
    fontWeight: 800,
    textTransform: "uppercase",
    letterSpacing: "0.06em",
    color: "#6b7280",
  },
  metricValue: {
    fontSize: 13,
    fontWeight: 800,
    color: "#0f172a",
  },
  section: {
    display: "grid",
    gap: 10,
    padding: 12,
    borderRadius: 14,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "rgba(255,255,255,0.9)",
    boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)",
  },
  sectionHeader: {
    display: "flex",
    justifyContent: "space-between",
    gap: 8,
    alignItems: "flex-start",
  },
  sectionTitle: {
    fontSize: 12,
    fontWeight: 900,
    textTransform: "uppercase",
    letterSpacing: "0.08em",
    color: "#0f172a",
  },
  sectionHint: {
    fontSize: 11,
    color: "#64748b",
    lineHeight: 1.4,
    marginTop: 2,
  },
  scrollBlock: {
    display: "grid",
    gap: 8,
    maxHeight: 260,
    overflowY: "auto",
    paddingRight: 4,
  },
  row: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: 8,
    alignItems: "center",
    padding: 10,
    borderRadius: 12,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "#fff",
  },
  rowActive: {
    display: "grid",
    gridTemplateColumns: "1fr auto",
    gap: 8,
    alignItems: "center",
    padding: 10,
    borderRadius: 12,
    border: "1px solid rgba(37, 99, 235, 0.35)",
    background: "rgba(219, 234, 254, 0.65)",
  },
  rowMain: {
    border: "none",
    background: "transparent",
    padding: 0,
    textAlign: "left",
    display: "grid",
    gap: 4,
    minWidth: 0,
    cursor: "pointer",
  },
  rowTitle: {
    fontSize: 13,
    fontWeight: 800,
    color: "#172b4d",
    overflow: "hidden",
    textOverflow: "ellipsis",
    whiteSpace: "nowrap",
  },
  rowMeta: {
    display: "flex",
    flexWrap: "wrap",
    gap: 8,
    fontSize: 11,
    color: "#6b778c",
  },
  rowActions: {
    display: "inline-flex",
    gap: 6,
    alignItems: "center",
  },
  iconBtn: {
    width: 30,
    height: 30,
    borderRadius: 999,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    background: "#fff",
    color: "#1d4ed8",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  iconBtnDanger: {
    width: 30,
    height: 30,
    borderRadius: 999,
    border: "1px solid rgba(239, 68, 68, 0.18)",
    background: "rgba(254, 226, 226, 0.9)",
    color: "#b91c1c",
    display: "inline-flex",
    alignItems: "center",
    justifyContent: "center",
    cursor: "pointer",
  },
  previewFrame: {
    borderRadius: 12,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    overflow: "hidden",
    background: "#f8fafc",
    minHeight: 260,
  },
  previewImage: {
    width: "100%",
    height: "100%",
    objectFit: "contain",
    display: "block",
    background: "#f8fafc",
  },
  previewIframe: {
    width: "100%",
    height: 420,
    border: "none",
    display: "block",
  },
  previewText: {
    margin: 0,
    padding: 12,
    background: "#f8fafc",
    borderRadius: 12,
    border: "1px solid rgba(15, 23, 42, 0.08)",
    fontFamily: "Consolas, monospace",
    fontSize: 12,
    lineHeight: 1.5,
    whiteSpace: "pre-wrap",
    maxHeight: 360,
    overflow: "auto",
  },
};
