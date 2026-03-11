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

type EmailSortMode = "date_desc" | "date_asc" | "subject_asc" | "subject_desc";
type EmailAttachmentFilter = "all" | "with" | "without";
type DocumentFilterMode = "all" | "selected_email";
type PreviewState =
  | { kind: "image"; dataUrl: string }
  | { kind: "pdf"; dataUrl: string }
  | { kind: "text"; text: string }
  | { kind: "unsupported" };

function closeExplorer() {
  try { (window as any).Office?.context?.ui?.messageParent?.("close"); } catch {}
  try { window.close(); } catch {}
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

function formatDate(value: string | undefined): string {
  const raw = String(value || "").trim();
  if (!raw) return "";
  const parsed = new Date(raw);
  if (Number.isNaN(parsed.getTime())) return raw;
  return parsed.toLocaleString("pt-PT", { day: "2-digit", month: "2-digit", hour: "2-digit", minute: "2-digit" });
}

function formatBytes(value: number | undefined): string {
  const size = Number(value || 0);
  if (!size) return "";
  if (size < 1024) return `${size} B`;
  if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
  return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function normalizeMessageId(value: string | undefined): string {
  return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function stripDataUrlPrefix(value: string | undefined): string {
  return String(value || "").trim().replace(/^data:[^,]+,/, "");
}

function makeEmailKey(email: Partial<RelatedEmailEntry>): string {
  return String(email?.id || email?.itemId || email?.internetMessageId || `${email?.conversationId || ""}|${email?.subject || ""}`);
}

function makeDocumentKey(document: Partial<GroupDocumentEntry>): string {
  return String(document?.id || document?.storagePathHint || document?.name || "");
}

function emailHasAttachments(email: RelatedEmailEntry): boolean {
  return Array.isArray(email.attachments) && email.attachments.length > 0;
}

function getEmailTimestamp(email: RelatedEmailEntry): number {
  const parsed = new Date(String(email.messageDateIso || email.receivedAtIso || email.sentAtIso || "").trim()).getTime();
  return Number.isFinite(parsed) ? parsed : 0;
}

function inferDocumentKind(document: GroupDocumentEntry): "image" | "pdf" | "text" | "unsupported" {
  const name = String(document.name || "").toLowerCase();
  const type = String(document.contentType || "").toLowerCase();
  if (!stripDataUrlPrefix(document.contentBase64)) return "unsupported";
  if (type.startsWith("image/") || /\.(png|jpe?g|gif|bmp|webp|svg)$/.test(name)) return "image";
  if (type.includes("pdf") || /\.pdf$/.test(name)) return "pdf";
  if (type.startsWith("text/") || type.includes("json") || type.includes("xml") || type.includes("csv") || /\.(txt|md|json|xml|csv|log|ya?ml)$/.test(name)) return "text";
  return "unsupported";
}

function buildDocumentPreview(document: GroupDocumentEntry | null): PreviewState | null {
  if (!document?.contentBase64) return null;
  const base64 = stripDataUrlPrefix(document.contentBase64);
  if (!base64) return null;
  const dataUrl = `data:${document.contentType || "application/octet-stream"};base64,${base64}`;
  const kind = inferDocumentKind(document);
  if (kind === "image") return { kind, dataUrl };
  if (kind === "pdf") return { kind, dataUrl };
  if (kind === "text") {
    try { return { kind, text: globalThis.atob(base64) }; } catch { return { kind: "unsupported" }; }
  }
  return { kind: "unsupported" };
}

function matchesSelectedEmail(document: GroupDocumentEntry, email: RelatedEmailEntry | null): boolean {
  if (!email) return false;
  const keys = new Set([
    String(email.id || "").trim(),
    String(email.itemId || "").trim(),
    normalizeMessageId(email.internetMessageId),
    String(email.conversationId || "").trim(),
    String(email.subject || "").trim(),
  ].filter(Boolean));
  return Boolean(
    (document.sourceEmailKey && keys.has(String(document.sourceEmailKey).trim()))
    || (document.sourceItemId && keys.has(String(document.sourceItemId).trim()))
    || (document.sourceInternetMessageId && keys.has(normalizeMessageId(document.sourceInternetMessageId)))
    || (document.sourceConversationId && keys.has(String(document.sourceConversationId).trim()))
    || (document.sourceEmailSubject && keys.has(String(document.sourceEmailSubject).trim()))
  );
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
  const [emailSearch, setEmailSearch] = useState("");
  const [emailAttachmentFilter, setEmailAttachmentFilter] = useState<EmailAttachmentFilter>("all");
  const [emailSort, setEmailSort] = useState<EmailSortMode>("date_desc");
  const [dateFrom, setDateFrom] = useState("");
  const [dateTo, setDateTo] = useState("");
  const [documentFilterMode, setDocumentFilterMode] = useState<DocumentFilterMode>("all");
  const [documentSearch, setDocumentSearch] = useState("");

  useEffect(() => {
    (async () => {
      try {
        const settings = await getSettings();
        if (settings.skinId) applySkin(settings.skinId);
      } catch {}
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
      .catch((nextError: any) => { if (!cancelled) setError(nextError?.message || "Nao foi possivel carregar os grupos."); })
      .finally(() => { if (!cancelled) setLoadingGroups(false); });
    return () => { cancelled = true; };
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
        setSelectedEmailKey((current) => current && emails.some((email) => makeEmailKey(email) === current) ? current : (emails[0] ? makeEmailKey(emails[0]) : ""));
        setSelectedDocumentId((current) => {
          if (current && documents.some((document) => makeDocumentKey(document) === current)) return current;
          if (initial.documentId && documents.some((document) => makeDocumentKey(document) === initial.documentId)) return initial.documentId;
          return "";
        });
      })
      .catch((nextError: any) => { if (!cancelled) setError(nextError?.message || "Nao foi possivel carregar os dados do grupo."); })
      .finally(() => {
        if (cancelled) return;
        setLoadingEmails(false);
        setLoadingDocuments(false);
      });
    return () => { cancelled = true; };
  }, [initial.documentId, initial.emailKey, selectedGroupId]);

  const selectedGroup = useMemo(() => groups.find((group) => group.id === selectedGroupId) || null, [groups, selectedGroupId]);

  const filteredEmails = useMemo(() => {
    const query = String(emailSearch || "").trim().toLowerCase();
    const fromTime = dateFrom ? new Date(`${dateFrom}T00:00:00`).getTime() : 0;
    const toTime = dateTo ? new Date(`${dateTo}T23:59:59`).getTime() : 0;
    const next = groupEmails.filter((email) => {
      if (query) {
        const haystack = [email.subject, email.fromName, email.fromEmail].map((value) => String(value || "").toLowerCase()).join(" ");
        if (!haystack.includes(query)) return false;
      }
      if (emailAttachmentFilter === "with" && !emailHasAttachments(email)) return false;
      if (emailAttachmentFilter === "without" && emailHasAttachments(email)) return false;
      if (fromTime || toTime) {
        const timestamp = getEmailTimestamp(email);
        if (fromTime && (!timestamp || timestamp < fromTime)) return false;
        if (toTime && (!timestamp || timestamp > toTime)) return false;
      }
      return true;
    });
    next.sort((a, b) => {
      if (emailSort === "date_asc") return getEmailTimestamp(a) - getEmailTimestamp(b);
      if (emailSort === "subject_asc") return String(a.subject || "").localeCompare(String(b.subject || ""), "pt");
      if (emailSort === "subject_desc") return String(b.subject || "").localeCompare(String(a.subject || ""), "pt");
      return getEmailTimestamp(b) - getEmailTimestamp(a);
    });
    return next;
  }, [dateFrom, dateTo, emailAttachmentFilter, emailSearch, emailSort, groupEmails]);

  const selectedEmail = useMemo(
    () => filteredEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || groupEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || filteredEmails[0] || groupEmails[0] || null,
    [filteredEmails, groupEmails, selectedEmailKey]
  );

  const filteredDocuments = useMemo(() => {
    const query = String(documentSearch || "").trim().toLowerCase();
    return groupDocuments.filter((document) => {
      if (documentFilterMode === "selected_email" && !matchesSelectedEmail(document, selectedEmail)) return false;
      if (!query) return true;
      const haystack = [document.name, document.sourceEmailSubject, document.contentType].map((value) => String(value || "").toLowerCase()).join(" ");
      return haystack.includes(query);
    });
  }, [documentFilterMode, documentSearch, groupDocuments, selectedEmail]);

  const selectedDocument = useMemo(
    () => filteredDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || groupDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || null,
    [filteredDocuments, groupDocuments, selectedDocumentId]
  );

  const selectedDocumentPreview = useMemo(() => buildDocumentPreview(selectedDocument), [selectedDocument]);
  const selectedProvider = useMemo(() => providerLabel(selectedDocument?.storageProvider || groupDocuments[0]?.storageProvider), [groupDocuments, selectedDocument]);

  useEffect(() => {
    if (!filteredEmails.some((email) => makeEmailKey(email) === selectedEmailKey)) {
      setSelectedEmailKey(filteredEmails[0] ? makeEmailKey(filteredEmails[0]) : "");
    }
  }, [filteredEmails, selectedEmailKey]);

  useEffect(() => {
    if (!filteredDocuments.some((document) => makeDocumentKey(document) === selectedDocumentId)) {
      setSelectedDocumentId("");
    }
  }, [filteredDocuments, selectedDocumentId]);

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
    const opened = await openLinkedOutlookEmail({ itemId: email.itemId, emailWebLink: email.emailWebLink });
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
    const base64 = stripDataUrlPrefix(document.contentBase64);
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
      await addBase64AttachmentToCompose(document.name || "documento", stripDataUrlPrefix(document.contentBase64));
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
          <div style={styles.subtitle}>Primeiro os documentos. Depois os emails ligados ao grupo, com filtros rapidos e preview mais visivel.</div>
        </div>
        <div style={styles.headerActions}>
          <div style={styles.selectWrap}>
            <select style={styles.select} value={selectedGroupId} onChange={(event) => { setSelectedGroupId(event.target.value); setSelectedEmailKey(""); setSelectedDocumentId(""); setNotice(null); }}>
              {groups.map((group) => <option key={group.id} value={group.id}>{group.name}</option>)}
            </select>
          </div>
          <button type="button" style={styles.iconBtn} onClick={() => void refreshCurrentGroup()} disabled={loadingGroups || loadingEmails || loadingDocuments}><Icons.RefreshCw size={14} /></button>
          <button type="button" style={styles.closeBtn} onClick={closeExplorer}>Fechar</button>
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
              <div style={styles.metric}><span style={styles.metricLabel}>Grupo</span><span style={styles.metricValue}>{selectedGroup.name}</span></div>
              <div style={styles.metric}><span style={styles.metricLabel}>Provider</span><span style={styles.metricValue}>{selectedProvider}</span></div>
              <div style={styles.metric}><span style={styles.metricLabel}>Documentos</span><span style={styles.metricValue}>{groupDocuments.length}</span></div>
              <div style={styles.metric}><span style={styles.metricLabel}>Emails</span><span style={styles.metricValue}>{groupEmails.length}</span></div>
            </div>
          </section>

          <section style={styles.columnsGrid}>
            <section style={styles.panel}>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={styles.sectionTitle}>Documentos</div>
                  <div style={styles.sectionHint}>Documentos do grupo na coluna principal, com filtro global ou por email selecionado.</div>
                </div>
              </div>
              <div style={styles.previewSection}>
                <div style={styles.sectionHeader}>
                  <div>
                    <div style={styles.sectionTitle}>Preview</div>
                    <div style={styles.sectionHint}>Seleciona um documento e o preview aparece logo aqui, sem esconder as listas.</div>
                  </div>
                  {selectedDocument ? (
                    <div style={styles.sectionActions}>
                      <button type="button" style={styles.iconBtn} onClick={() => handleDownloadDocument(selectedDocument)} disabled={!selectedDocument.contentBase64}><Icons.Download size={13} /></button>
                      <button type="button" style={styles.iconBtn} onClick={() => void handleAttachDocument(selectedDocument)} disabled={!selectedDocument.contentBase64}><Icons.Upload size={13} /></button>
                    </div>
                  ) : null}
                </div>
                {!selectedDocument ? <PanelState compact tone="info" title="Sem documento selecionado" description="Escolhe um documento desta coluna para abrir o preview." /> : null}
                {selectedDocument && selectedDocumentPreview?.kind === "image" ? <div style={styles.previewFrame}><img src={selectedDocumentPreview.dataUrl} alt={selectedDocument.name} style={styles.previewImage} /></div> : null}
                {selectedDocument && selectedDocumentPreview?.kind === "pdf" ? <div style={styles.previewFrame}><iframe title={selectedDocument.name} src={selectedDocumentPreview.dataUrl} style={styles.previewIframe} /></div> : null}
                {selectedDocument && selectedDocumentPreview?.kind === "text" ? <pre style={styles.previewText}>{selectedDocumentPreview.text}</pre> : null}
                {selectedDocument && (!selectedDocumentPreview || selectedDocumentPreview.kind === "unsupported") ? <PanelState compact tone="info" title="Preview nao disponivel" description="Este documento pode ser descarregado ou anexado, mas ainda nao tem preview interno para este formato." /> : null}
              </div>
              <div style={styles.filterGrid}>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Filtro</span>
                  <select style={styles.compactSelect} value={documentFilterMode} onChange={(event) => setDocumentFilterMode(event.target.value as DocumentFilterMode)}>
                    <option value="all">Todos os documentos</option>
                    <option value="selected_email">Do email selecionado</option>
                  </select>
                </label>
                <label style={styles.filterFieldWide}>
                  <span style={styles.filterLabel}>Pesquisar documento</span>
                  <input style={styles.input} value={documentSearch} onChange={(event) => setDocumentSearch(event.target.value)} placeholder="Nome, tipo ou assunto..." />
                </label>
              </div>
              <div style={styles.listShell}>
                {loadingDocuments && !groupDocuments.length ? <PanelState compact tone="loading" title="A carregar documentos" description="A listar os documentos guardados." /> : null}
                {!loadingDocuments && !filteredDocuments.length ? <PanelState compact tone="info" title="Sem documentos visiveis" description={documentFilterMode === "selected_email" ? "Nao ha documentos associados ao email atualmente selecionado." : "Este grupo ainda nao tem documentos guardados visiveis neste filtro."} /> : null}
                {filteredDocuments.map((document) => {
                  const active = makeDocumentKey(document) === makeDocumentKey(selectedDocument || {});
                  return (
                    <div key={makeDocumentKey(document)} style={active ? styles.cardActive : styles.card}>
                      <button type="button" style={styles.cardMain} onClick={() => setSelectedDocumentId(makeDocumentKey(document))}>
                        <div style={styles.cardTitle}>{document.name}</div>
                        <div style={styles.cardMeta}>
                          <span>{document.contentType || "Documento"}</span>
                          {formatBytes(document.size) ? <span>{formatBytes(document.size)}</span> : null}
                          {document.sourceEmailSubject ? <span>{document.sourceEmailSubject}</span> : null}
                        </div>
                      </button>
                      <div style={styles.cardActions}>
                        <button type="button" style={styles.iconBtn} onClick={() => handleDownloadDocument(document)} disabled={!document.contentBase64}><Icons.Download size={12} /></button>
                        <button type="button" style={styles.iconBtn} onClick={() => void handleAttachDocument(document)} disabled={!document.contentBase64}><Icons.Upload size={12} /></button>
                        <button type="button" style={styles.iconBtnDanger} onClick={() => void handleDeleteDocument(document)} disabled={busy}><Icons.Trash size={12} /></button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>

            <section style={styles.panel}>
              <div style={styles.sectionHeader}>
                <div>
                  <div style={styles.sectionTitle}>Emails</div>
                  <div style={styles.sectionHint}>Emails ligados ao grupo, com pesquisa, filtros rapidos e ordenacao.</div>
                </div>
              </div>
              <div style={styles.filterGrid}>
                <label style={styles.filterFieldWide}>
                  <span style={styles.filterLabel}>Pesquisar email</span>
                  <input style={styles.input} value={emailSearch} onChange={(event) => setEmailSearch(event.target.value)} placeholder="Assunto, contacto ou email..." />
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Anexos</span>
                  <select style={styles.compactSelect} value={emailAttachmentFilter} onChange={(event) => setEmailAttachmentFilter(event.target.value as EmailAttachmentFilter)}>
                    <option value="all">Todos</option>
                    <option value="with">Com anexos</option>
                    <option value="without">Sem anexos</option>
                  </select>
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Ordenar</span>
                  <select style={styles.compactSelect} value={emailSort} onChange={(event) => setEmailSort(event.target.value as EmailSortMode)}>
                    <option value="date_desc">Mais recentes</option>
                    <option value="date_asc">Mais antigos</option>
                    <option value="subject_asc">A-Z</option>
                    <option value="subject_desc">Z-A</option>
                  </select>
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>De</span>
                  <input style={styles.compactInput} type="date" value={dateFrom} onChange={(event) => setDateFrom(event.target.value)} />
                </label>
                <label style={styles.filterField}>
                  <span style={styles.filterLabel}>Ate</span>
                  <input style={styles.compactInput} type="date" value={dateTo} onChange={(event) => setDateTo(event.target.value)} />
                </label>
              </div>
              <div style={styles.listShell}>
                {loadingEmails && !groupEmails.length ? <PanelState compact tone="loading" title="A carregar emails" description="A listar os emails do grupo." /> : null}
                {!loadingEmails && !filteredEmails.length ? <PanelState compact tone="info" title="Sem emails visiveis" description="Nao ha emails a corresponder aos filtros atuais." /> : null}
                {filteredEmails.map((email) => {
                  const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                  const canOpen = Boolean(email.itemId || email.emailWebLink);
                  return (
                    <div key={makeEmailKey(email)} style={active ? styles.cardActive : styles.card}>
                      <button type="button" style={styles.cardMain} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                        <div style={styles.cardTitle}>{email.subject || "(sem assunto)"}</div>
                        <div style={styles.cardMeta}>
                          <span>{email.fromName || email.fromEmail || "(sem remetente)"}</span>
                          {formatDate(email.messageDateIso || email.receivedAtIso) ? <span>{formatDate(email.messageDateIso || email.receivedAtIso)}</span> : null}
                          <span>{emailHasAttachments(email) ? `${email.attachments?.length || 0} anexo(s)` : "sem anexos"}</span>
                        </div>
                      </button>
                      <div style={styles.cardActions}>
                        <button type="button" style={styles.iconBtn} onClick={() => void handleOpenEmail(email)} disabled={!canOpen}><Icons.MessageSquare size={12} /></button>
                        <button type="button" style={styles.iconBtnDanger} onClick={() => void handleRemoveEmail(email)} disabled={busy}><Icons.Trash size={12} /></button>
                      </div>
                    </div>
                  );
                })}
              </div>
            </section>
          </section>
        </>
      ) : null}
    </div>
  );
}

const styles: Record<string, React.CSSProperties> = {
  root: { minHeight: "100vh", background: "var(--iccc-bg, #edf2f7)", color: "var(--iccc-text, #172b4d)", fontFamily: "var(--iccc-font, 'Segoe UI', sans-serif)", padding: 14, display: "grid", gap: 12 },
  header: { display: "grid", gap: 10, padding: 12, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.86)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.08)" },
  headerCopy: { display: "grid", gap: 4 },
  eyebrow: { fontSize: 11, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#5b6b83" },
  title: { fontSize: 24, fontWeight: 800, color: "#0f172a" },
  subtitle: { fontSize: 12, color: "#607086", lineHeight: 1.45 },
  headerActions: { display: "flex", gap: 8, alignItems: "center", flexWrap: "wrap" },
  selectWrap: { flex: "1 1 240px", minWidth: 180 },
  select: { width: "100%", borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.12)", background: "rgba(248,250,252,0.95)", color: "#172b4d", padding: "9px 14px", fontSize: 12, fontWeight: 600, outline: "none" },
  closeBtn: { borderRadius: 999, border: "none", background: "#1d4ed8", color: "#fff", padding: "9px 16px", fontSize: 12, fontWeight: 700, cursor: "pointer" },
  summaryCard: { padding: 12, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.86)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)" },
  metricGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))", gap: 8 },
  metric: { display: "grid", gap: 2, padding: 10, borderRadius: 12, background: "rgba(15, 23, 42, 0.03)" },
  metricLabel: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.06em", color: "#6b7280" },
  metricValue: { fontSize: 13, fontWeight: 700, color: "#0f172a" },
  previewSection: { display: "grid", gap: 10, padding: 10, borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(248,250,252,0.85)" },
  columnsGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(320px, 1fr))", gap: 12, alignItems: "start" },
  panel: { display: "grid", gap: 10, padding: 12, borderRadius: 14, border: "1px solid rgba(15, 23, 42, 0.08)", background: "rgba(255,255,255,0.9)", boxShadow: "0 10px 28px rgba(15, 23, 42, 0.05)", minHeight: 0 },
  sectionHeader: { display: "flex", justifyContent: "space-between", gap: 8, alignItems: "flex-start" },
  sectionTitle: { fontSize: 12, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.08em", color: "#0f172a" },
  sectionHint: { fontSize: 11, color: "#64748b", lineHeight: 1.4, marginTop: 2 },
  sectionActions: { display: "inline-flex", gap: 6, alignItems: "center" },
  filterGrid: { display: "grid", gridTemplateColumns: "repeat(auto-fit, minmax(120px, 1fr))", gap: 8, alignItems: "end" },
  filterField: { display: "grid", gap: 4, minWidth: 0 },
  filterFieldWide: { display: "grid", gap: 4, minWidth: 0, gridColumn: "span 2" },
  filterLabel: { fontSize: 10, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.05em", color: "#64748b" },
  input: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "8px 10px", fontSize: 12, outline: "none", minWidth: 0 },
  compactInput: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "8px 10px", fontSize: 12, outline: "none", minWidth: 0 },
  compactSelect: { width: "100%", borderRadius: 10, border: "1px solid rgba(15, 23, 42, 0.12)", background: "#f8fafc", color: "#172b4d", padding: "8px 10px", fontSize: 12, outline: "none", minWidth: 0 },
  listShell: { display: "grid", gap: 8, maxHeight: "36vh", overflowY: "auto", paddingRight: 4 },
  card: { display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "center", padding: 9, borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff" },
  cardActive: { display: "grid", gridTemplateColumns: "1fr auto", gap: 8, alignItems: "center", padding: 9, borderRadius: 12, border: "1px solid rgba(37, 99, 235, 0.35)", background: "rgba(219, 234, 254, 0.65)" },
  cardMain: { border: "none", background: "transparent", padding: 0, textAlign: "left", display: "grid", gap: 4, minWidth: 0, cursor: "pointer" },
  cardTitle: { fontSize: 12, fontWeight: 600, color: "#172b4d", lineHeight: 1.35, wordBreak: "break-word" },
  cardMeta: { display: "flex", flexWrap: "wrap", gap: 8, fontSize: 10.5, color: "#6b778c", lineHeight: 1.35 },
  cardActions: { display: "inline-flex", gap: 6, alignItems: "center" },
  iconBtn: { width: 30, height: 30, borderRadius: 999, border: "1px solid rgba(15, 23, 42, 0.08)", background: "#fff", color: "#1d4ed8", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  iconBtnDanger: { width: 30, height: 30, borderRadius: 999, border: "1px solid rgba(239, 68, 68, 0.18)", background: "rgba(254, 226, 226, 0.9)", color: "#b91c1c", display: "inline-flex", alignItems: "center", justifyContent: "center", cursor: "pointer" },
  previewFrame: { borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", overflow: "hidden", background: "#f8fafc", minHeight: 220 },
  previewImage: { width: "100%", height: "100%", minHeight: 220, maxHeight: 300, objectFit: "contain", display: "block", background: "#fff" },
  previewIframe: { width: "100%", height: 300, border: "none", display: "block", background: "#fff" },
  previewText: { margin: 0, padding: 12, background: "#f8fafc", borderRadius: 12, border: "1px solid rgba(15, 23, 42, 0.08)", fontFamily: "Consolas, monospace", fontSize: 12, lineHeight: 1.5, whiteSpace: "pre-wrap", maxHeight: 280, overflow: "auto" },
};
