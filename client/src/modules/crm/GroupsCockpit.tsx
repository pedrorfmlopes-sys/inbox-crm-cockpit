import React, { useEffect, useMemo, useRef, useState } from "react";
import {
    addEmailToLinkGroup,
    createLinkGroup,
    deleteLinkGroup,
    deleteGroupDocument,
    getGroupEmails,
    getGroupDocuments,
    listLinkGroups,
    removeEmailFromLinkGroup,
    saveGroupDocuments,
    type GroupDocumentEntry,
    type LinkGroupEntry,
    type RelatedEmailEntry,
} from "@/api";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { addBase64AttachmentToCompose, openLinkedOutlookEmail } from "@/office";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type CurrentAttachmentCandidate = {
    id: string;
    name: string;
    contentType?: string;
    content: string;
    size?: number;
};

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

function estimateBase64Size(base64: string | undefined): number {
    const raw = String(base64 || "").trim().replace(/^data:[^,]+,/, "");
    if (!raw) return 0;
    const padding = raw.endsWith("==") ? 2 : raw.endsWith("=") ? 1 : 0;
    return Math.max(0, Math.floor((raw.length * 3) / 4) - padding);
}

function formatBytes(value: number | undefined): string {
    const size = Number(value || 0);
    if (!size) return "";
    if (size < 1024) return `${size} B`;
    if (size < 1024 * 1024) return `${(size / 1024).toFixed(1)} KB`;
    return `${(size / (1024 * 1024)).toFixed(1)} MB`;
}

function sanitizePathSegment(value: string): string {
    return String(value || "").trim().replace(/[\\/:*?"<>|]+/g, "_");
}

function IconButton({
    title,
    icon,
    onClick,
    disabled,
    tone = "default",
}: {
    title: string;
    icon: React.ReactNode;
    onClick?: () => void;
    disabled?: boolean;
    tone?: "default" | "primary" | "danger";
}) {
    const style =
        tone === "primary"
            ? styles.iconBtnPrimary
            : tone === "danger"
                ? styles.iconBtnDanger
                : styles.iconBtn;
    return (
        <button type="button" title={title} aria-label={title} style={style} onClick={onClick} disabled={disabled}>
            {icon}
        </button>
    );
}

function Section({
    title,
    subtitle,
    actions,
    children,
}: {
    title: string;
    subtitle?: string;
    actions?: React.ReactNode;
    children: React.ReactNode;
}) {
    return (
        <section style={styles.section}>
            <div style={styles.sectionHeader}>
                <div style={styles.sectionTitleWrap}>
                    <div style={styles.sectionTitle}>{title}</div>
                    {subtitle ? <div style={styles.sectionSubtitle}>{subtitle}</div> : null}
                </div>
                {actions ? <div style={styles.sectionActions}>{actions}</div> : null}
            </div>
            <div style={styles.sectionBody}>{children}</div>
        </section>
    );
}

export const GroupsCockpit: React.FC = () => {
    const { ctx, attachments, setMsg, settings } = useCockpit();
    const downloadAnchorRef = useRef<HTMLAnchorElement | null>(null);
    const [query, setQuery] = useState("");
    const [newGroupName, setNewGroupName] = useState("");
    const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
    const [selectedGroupId, setSelectedGroupId] = useState("");
    const [selectedEmailKey, setSelectedEmailKey] = useState("");
    const [selectedDocumentId, setSelectedDocumentId] = useState("");
    const [groupEmails, setGroupEmails] = useState<RelatedEmailEntry[]>([]);
    const [groupDocuments, setGroupDocuments] = useState<GroupDocumentEntry[]>([]);
    const [groupsLoading, setGroupsLoading] = useState(false);
    const [emailsLoading, setEmailsLoading] = useState(false);
    const [documentsLoading, setDocumentsLoading] = useState(false);
    const [groupsError, setGroupsError] = useState<string | null>(null);
    const [emailsError, setEmailsError] = useState<string | null>(null);
    const [documentsError, setDocumentsError] = useState<string | null>(null);
    const [busyAction, setBusyAction] = useState(false);
    const [reloadToken, setReloadToken] = useState(0);

    const selectedGroup = useMemo(
        () => groups.find((group) => group.id === selectedGroupId) || null,
        [groups, selectedGroupId]
    );

    const currentEmailPayload = useMemo(
        () => ({
            itemId: String(ctx.itemId || "").trim(),
            internetMessageId: String(ctx.internetMessageId || "").trim(),
            conversationId: String(ctx.conversationId || "").trim(),
            subject: String(ctx.subject || "").trim(),
            fromEmail: String(ctx.fromEmail || "").trim(),
            fromName: String(ctx.fromName || "").trim(),
            receivedAtIso: String(ctx.receivedDateTimeIso || "").trim(),
            messageDateIso: String(ctx.receivedDateTimeIso || "").trim(),
            attachments: (attachments || []).map((attachment) => ({
                name: attachment.name,
                contentType: attachment.contentType,
            })),
        }),
        [attachments, ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject]
    );

    const hasCurrentEmail = Boolean(
        currentEmailPayload.itemId || currentEmailPayload.internetMessageId || currentEmailPayload.conversationId
    );

    useEffect(() => {
        let cancelled = false;
        const timer = window.setTimeout(async () => {
            setGroupsLoading(true);
            setGroupsError(null);
            try {
                const nextGroups = await listLinkGroups(query);
                if (cancelled) return;
                setGroups(nextGroups);
                setSelectedGroupId((current) => {
                    if (current && nextGroups.some((group) => group.id === current)) return current;
                    return nextGroups[0]?.id || "";
                });
            } catch (error: any) {
                if (cancelled) return;
                setGroupsError(error?.message || "Nao foi possivel carregar grupos.");
                setGroups([]);
                setSelectedGroupId("");
            } finally {
                if (!cancelled) setGroupsLoading(false);
            }
        }, 180);
        return () => {
            cancelled = true;
            window.clearTimeout(timer);
        };
    }, [query, reloadToken]);

    useEffect(() => {
        if (!selectedGroupId) {
            setGroupEmails([]);
            setSelectedEmailKey("");
            return;
        }
        let cancelled = false;
        setEmailsLoading(true);
        setEmailsError(null);
        getGroupEmails(selectedGroupId)
            .then((emails) => {
                if (cancelled) return;
                setGroupEmails(emails);
                setSelectedEmailKey((current) => {
                    if (current && emails.some((email) => makeEmailKey(email) === current)) return current;
                    return emails[0] ? makeEmailKey(emails[0]) : "";
                });
            })
            .catch((error: any) => {
                if (cancelled) return;
                setEmailsError(error?.message || "Nao foi possivel carregar os emails do grupo.");
                setGroupEmails([]);
                setSelectedEmailKey("");
            })
            .finally(() => {
                if (!cancelled) setEmailsLoading(false);
            });

        return () => {
            cancelled = true;
        };
    }, [selectedGroupId, reloadToken]);

    useEffect(() => {
        if (!selectedGroupId) {
            setGroupDocuments([]);
            setSelectedDocumentId("");
            return;
        }
        let cancelled = false;
        setDocumentsLoading(true);
        setDocumentsError(null);
        getGroupDocuments(selectedGroupId)
            .then((documents) => {
                if (cancelled) return;
                setGroupDocuments(documents);
                setSelectedDocumentId((current) => {
                    if (current && documents.some((document) => makeDocumentKey(document) === current)) return current;
                    return documents[0] ? makeDocumentKey(documents[0]) : "";
                });
            })
            .catch((error: any) => {
                if (cancelled) return;
                setDocumentsError(error?.message || "Nao foi possivel carregar os documentos do grupo.");
                setGroupDocuments([]);
                setSelectedDocumentId("");
            })
            .finally(() => {
                if (!cancelled) setDocumentsLoading(false);
            });

        return () => {
            cancelled = true;
        };
    }, [selectedGroupId, reloadToken]);

    const selectedEmail = useMemo(
        () => groupEmails.find((email) => makeEmailKey(email) === selectedEmailKey) || groupEmails[0] || null,
        [groupEmails, selectedEmailKey]
    );

    const selectedDocument = useMemo(
        () => groupDocuments.find((document) => makeDocumentKey(document) === selectedDocumentId) || groupDocuments[0] || null,
        [groupDocuments, selectedDocumentId]
    );

    const currentAttachmentCandidates = useMemo<CurrentAttachmentCandidate[]>(() => {
        const ignoreInline = Boolean(settings?.groupStorage.ignoreInlineAttachments);
        return (attachments || [])
            .map((attachment, index) => ({
                id: `${attachment.name}:${index}`,
                name: String(attachment.name || "").trim(),
                contentType: String(attachment.contentType || "").trim(),
                content: String(attachment.content || "").trim(),
                size: estimateBase64Size(attachment.content),
            }))
            .filter((attachment) => attachment.name && attachment.content)
            .filter((attachment) => {
                if (!ignoreInline) return true;
                const lowerName = attachment.name.toLowerCase();
                const lowerType = String(attachment.contentType || "").toLowerCase();
                if (lowerType.startsWith("image/") && (/^image\d+\./.test(lowerName) || lowerName.includes("logo") || lowerName.includes("signature") || lowerName.includes("assinatura"))) {
                    return false;
                }
                return true;
            });
    }, [attachments, settings?.groupStorage.ignoreInlineAttachments]);

    const groupFolderHint = useMemo(() => {
        const base = String(settings?.groupStorage.baseFolderPath || "").trim();
        const groupName = String(selectedGroup?.name || "").trim();
        if (!groupName) return "";
        if (!base) return sanitizePathSegment(groupName);
        const separator = /^https?:\/\//i.test(base) || base.endsWith("/") ? "/" : base.includes("\\") ? "\\" : "/";
        return `${base.replace(/[\\/]+$/, "")}${separator}${sanitizePathSegment(groupName)}`;
    }, [selectedGroup?.name, settings?.groupStorage.baseFolderPath]);

    const selectedDocumentPreview = useMemo(() => {
        if (!selectedDocument?.contentBase64) return null;
        const contentType = String(selectedDocument.contentType || "").toLowerCase();
        const dataUrl = `data:${selectedDocument.contentType || "application/octet-stream"};base64,${selectedDocument.contentBase64}`;
        if (contentType.startsWith("image/")) {
            return { kind: "image" as const, dataUrl };
        }
        if (contentType === "application/pdf") {
            return { kind: "pdf" as const, dataUrl };
        }
        if (contentType.startsWith("text/") || contentType.includes("json") || contentType.includes("xml")) {
            try {
                return { kind: "text" as const, text: globalThis.atob(selectedDocument.contentBase64) };
            } catch {
                return { kind: "unsupported" as const };
            }
        }
        return { kind: "unsupported" as const };
    }, [selectedDocument]);

    async function refreshGroupsAndEmails() {
        setReloadToken((value) => value + 1);
    }

    async function handleCreateGroup() {
        const name = String(newGroupName || "").trim();
        if (!name) return;
        setBusyAction(true);
        try {
            const group = await createLinkGroup({ name });
            setNewGroupName("");
            setSelectedGroupId(group.id);
            setReloadToken((value) => value + 1);
            setMsg(`Grupo "${group.name}" criado.`);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel criar o grupo.");
        } finally {
            setBusyAction(false);
        }
    }

    async function handleLinkCurrentEmail() {
        if (!selectedGroup || !hasCurrentEmail) return;
        setBusyAction(true);
        try {
            await addEmailToLinkGroup(selectedGroup.id, currentEmailPayload);
            setReloadToken((value) => value + 1);
            setMsg(`Email atual associado ao grupo "${selectedGroup.name}".`);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel associar o email ao grupo.");
        } finally {
            setBusyAction(false);
        }
    }

    async function handleRemoveEmail(email: RelatedEmailEntry) {
        if (!selectedGroup) return;
        setBusyAction(true);
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
            setReloadToken((value) => value + 1);
            setMsg("Email removido do grupo.");
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel remover o email do grupo.");
        } finally {
            setBusyAction(false);
        }
    }

    async function handleOpenEmail(email: RelatedEmailEntry) {
        const opened = await openLinkedOutlookEmail({
            itemId: email.itemId,
            emailWebLink: email.emailWebLink,
        });
        if (!opened) {
            setMsg("Este email ainda nao tem abertura direta disponivel.");
        }
    }

    async function handleSaveAttachmentsToGroup(candidates: CurrentAttachmentCandidate[]) {
        if (!selectedGroup || !candidates.length) return;
        setBusyAction(true);
        try {
            const storageProvider = String(settings?.groupStorage.provider || "").trim();
            const storageBasePath = String(settings?.groupStorage.baseFolderPath || "").trim();
            const payloadDocs: GroupDocumentEntry[] = candidates.map((attachment) => ({
                id: `doc_${globalThis.crypto?.randomUUID?.() || `${Date.now()}_${attachment.id}`}`,
                name: attachment.name,
                contentType: attachment.contentType,
                contentBase64: attachment.content,
                size: attachment.size,
                sourceEmailKey: String(ctx.itemId || ctx.internetMessageId || ctx.conversationId || "").trim(),
                sourceItemId: String(ctx.itemId || "").trim(),
                sourceInternetMessageId: String(ctx.internetMessageId || "").trim(),
                sourceConversationId: String(ctx.conversationId || "").trim(),
                sourceEmailSubject: String(ctx.subject || "").trim(),
                storageProvider,
                storageBasePath,
                storagePathHint: groupFolderHint ? `${groupFolderHint}${groupFolderHint.includes("\\") ? "\\" : "/"}${sanitizePathSegment(attachment.name)}` : "",
            }));
            const result = await saveGroupDocuments(selectedGroup.id, { documents: payloadDocs });
            setGroupDocuments(result.documents || []);
            setSelectedDocumentId(result.documents?.[0] ? makeDocumentKey(result.documents[0]) : "");
            setMsg(candidates.length === 1 ? `Documento "${candidates[0].name}" guardado no grupo.` : `${candidates.length} documento(s) guardados no grupo.`);
            setReloadToken((value) => value + 1);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel guardar os anexos neste grupo.");
        } finally {
            setBusyAction(false);
        }
    }

    async function handleDeleteDocument(document: GroupDocumentEntry) {
        if (!selectedGroup || !document?.id) return;
        setBusyAction(true);
        try {
            await deleteGroupDocument(selectedGroup.id, document.id);
            setGroupDocuments((current) => current.filter((entry) => entry.id !== document.id));
            setSelectedDocumentId((current) => (current === makeDocumentKey(document) ? "" : current));
            setMsg(`Documento "${document.name}" removido do grupo.`);
            setReloadToken((value) => value + 1);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel remover o documento.");
        } finally {
            setBusyAction(false);
        }
    }

    function handleDownloadDocument(doc: GroupDocumentEntry) {
        const base64 = String(doc?.contentBase64 || "").trim();
        if (!base64) {
            setMsg("Este documento nao tem conteudo disponivel para download.");
            return;
        }
        const byteCharacters = globalThis.atob(base64);
        const byteNumbers = new Array(byteCharacters.length);
        for (let i = 0; i < byteCharacters.length; i += 1) byteNumbers[i] = byteCharacters.charCodeAt(i);
        const blob = new Blob([new Uint8Array(byteNumbers)], { type: doc.contentType || "application/octet-stream" });
        const url = URL.createObjectURL(blob);
        const anchor = downloadAnchorRef.current || globalThis.document.createElement("a");
        downloadAnchorRef.current = anchor;
        anchor.href = url;
        anchor.download = doc.name || "documento";
        anchor.click();
        setTimeout(() => URL.revokeObjectURL(url), 2000);
    }

    async function handleAttachDocument(document: GroupDocumentEntry) {
        try {
            await addBase64AttachmentToCompose(document.name || "documento", String(document.contentBase64 || ""));
            setMsg(`Documento "${document.name}" anexado ao email em edicao.`);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel anexar o documento.");
        }
    }

    return (
        <div style={styles.root}>
            <Section
                title="Grupos"
                subtitle="Pesquisa, cria e seleciona grupos manuais do Cockpit."
                actions={
                    <>
                        <IconButton
                            title={selectedGroup ? "Apagar grupo selecionado" : "Seleciona um grupo para apagar"}
                            icon={<Icons.Trash size={13} />}
                            onClick={() => {
                                if (!selectedGroup || busyAction) return;
                                const confirmed = window.confirm(`Apagar o grupo "${selectedGroup.name}"?`);
                                if (!confirmed) return;
                                void (async () => {
                                    setBusyAction(true);
                                    try {
                                        await deleteLinkGroup(selectedGroup.id);
                                        setSelectedGroupId("");
                                        setGroupEmails([]);
                                        setSelectedEmailKey("");
                                        setReloadToken((value) => value + 1);
                                        setMsg(`Grupo "${selectedGroup.name}" apagado.`);
                                    } catch (error: any) {
                                        setMsg(error?.message || "Nao foi possivel apagar o grupo.");
                                    } finally {
                                        setBusyAction(false);
                                    }
                                })();
                            }}
                            disabled={!selectedGroup || busyAction}
                            tone="danger"
                        />
                        <IconButton title="Atualizar grupos" icon={<Icons.RefreshCw size={13} />} onClick={refreshGroupsAndEmails} disabled={groupsLoading || busyAction} />
                    </>
                }
            >
                <div style={styles.inputStack}>
                    <div style={styles.compactField}>
                        <input
                            style={styles.input}
                            value={query}
                            onChange={(event) => setQuery(event.target.value)}
                            placeholder="Pesquisar grupos..."
                        />
                    </div>
                    <div style={styles.compactFieldRow}>
                        <input
                            style={styles.input}
                            value={newGroupName}
                            onChange={(event) => setNewGroupName(event.target.value)}
                            onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                    event.preventDefault();
                                    void handleCreateGroup();
                                }
                            }}
                            placeholder="Novo grupo..."
                        />
                        <IconButton title="Criar grupo" icon={<Icons.Plus size={13} />} onClick={handleCreateGroup} disabled={busyAction || !newGroupName.trim()} tone="primary" />
                    </div>
                </div>

                {groupsError ? <div style={styles.errorText}>{groupsError}</div> : null}

                <div style={styles.scrollPaneTop}>
                    {groupsLoading && !groups.length ? <PanelState compact tone="info" title="A carregar grupos" description="A sincronizar grupos guardados." /> : null}
                    {!groupsLoading && !groups.length ? <PanelState compact tone="info" title="Sem grupos" description="Cria o primeiro grupo manual para comecar a organizar emails." /> : null}
                    {groups.map((group) => {
                        const selected = group.id === selectedGroupId;
                        return (
                            <button key={group.id} type="button" style={selected ? styles.groupRowActive : styles.groupRow} onClick={() => setSelectedGroupId(group.id)}>
                                <span style={styles.groupMain}>
                                    <span style={styles.groupName}>{group.name}</span>
                                    {group.description ? <span style={styles.groupDesc}>{group.description}</span> : null}
                                </span>
                                <span style={styles.groupCount}>{group.memberCount || 0}</span>
                            </button>
                        );
                    })}
                </div>
            </Section>

            <Section
                title="Emails"
                subtitle={selectedGroup ? `Grupo selecionado: ${selectedGroup.name}` : "Seleciona um grupo para ver os emails associados."}
                actions={
                    <>
                        <IconButton
                            title={selectedGroup && hasCurrentEmail ? "Associar email aberto ao grupo" : "Abre um email para o associares ao grupo"}
                            icon={<Icons.Link size={13} />}
                            onClick={handleLinkCurrentEmail}
                            disabled={!selectedGroup || !hasCurrentEmail || busyAction}
                            tone="primary"
                        />
                        <IconButton
                            title="Atualizar emails do grupo"
                            icon={<Icons.RefreshCw size={13} />}
                            onClick={refreshGroupsAndEmails}
                            disabled={!selectedGroup || emailsLoading || busyAction}
                        />
                    </>
                }
            >
                {!hasCurrentEmail ? <div style={styles.hintText}>A associacao continua a ser feita a partir do email atualmente aberto no add-in.</div> : null}
                {emailsError ? <div style={styles.errorText}>{emailsError}</div> : null}
                <div style={styles.scrollPaneMiddle}>
                    {!selectedGroup ? <PanelState compact tone="info" title="Nenhum grupo selecionado" description="Escolhe um grupo acima para ver ou gerir os emails." /> : null}
                    {selectedGroup && emailsLoading && !groupEmails.length ? <PanelState compact tone="info" title="A carregar emails" description="A listar os emails ligados ao grupo." /> : null}
                    {selectedGroup && !emailsLoading && !groupEmails.length ? <PanelState compact tone="info" title="Grupo sem emails" description="Podes associar o email atualmente aberto com o botao de ligacao." /> : null}
                    {groupEmails.map((email) => {
                        const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                        const attachmentCount = Array.isArray(email.attachments) ? email.attachments.length : 0;
                        const canOpen = Boolean(email.itemId || email.emailWebLink);
                        return (
                            <div key={makeEmailKey(email)} style={active ? styles.emailRowActive : styles.emailRow}>
                                <button type="button" style={styles.emailSelectArea} onClick={() => setSelectedEmailKey(makeEmailKey(email))}>
                                    <div style={styles.emailSubject}>{email.subject || "(sem assunto)"}</div>
                                    <div style={styles.emailMeta}>
                                        <span>{email.fromName || email.fromEmail || "(sem remetente)"}</span>
                                        {formatDate(email.messageDateIso || email.receivedAtIso) ? <span>{formatDate(email.messageDateIso || email.receivedAtIso)}</span> : null}
                                    </div>
                                    <div style={styles.emailTagRow}>
                                        {attachmentCount ? <span style={styles.metaTag}>{attachmentCount} anexo(s)</span> : null}
                                        {Array.isArray(email.relatedRecords) && email.relatedRecords.length ? (
                                            <span style={styles.metaTag}>{email.relatedRecords.length} registo(s) Odoo</span>
                                        ) : null}
                                    </div>
                                </button>
                                <div style={styles.emailActions}>
                                    <IconButton title={canOpen ? "Abrir email" : "Sem abertura direta disponivel"} icon={<Icons.MessageSquare size={12} />} onClick={canOpen ? () => void handleOpenEmail(email) : undefined} disabled={!canOpen} />
                                    <IconButton title="Remover do grupo" icon={<Icons.Trash size={12} />} onClick={() => void handleRemoveEmail(email)} disabled={busyAction} tone="danger" />
                                </div>
                            </div>
                        );
                    })}
                </div>
            </Section>

            <Section
                title="Documentos e acoes"
                subtitle={selectedGroup ? "Documentos guardados no grupo e anexos disponiveis do email aberto." : "Seleciona um grupo para começar a gerir documentos."}
                actions={
                    <>
                        <IconButton
                            title={selectedGroup && currentAttachmentCandidates.length ? "Guardar todos os anexos do email aberto no grupo" : "Abre um email com anexos para os guardar no grupo"}
                            icon={<Icons.Save size={13} />}
                            onClick={selectedGroup && currentAttachmentCandidates.length ? () => void handleSaveAttachmentsToGroup(currentAttachmentCandidates) : undefined}
                            disabled={!selectedGroup || !currentAttachmentCandidates.length || busyAction}
                            tone="primary"
                        />
                        <IconButton
                            title={selectedGroup ? "Atualizar documentos do grupo" : "Seleciona um grupo"}
                            icon={<Icons.RefreshCw size={13} />}
                            onClick={selectedGroup ? refreshGroupsAndEmails : undefined}
                            disabled={!selectedGroup || documentsLoading || busyAction}
                        />
                    </>
                }
            >
                <div style={styles.actionStrip}>
                    <span style={styles.actionHint}>
                        {selectedGroup ? `${groupDocuments.length} documento(s) guardado(s)` : "Sem grupo ativo"}
                    </span>
                    {selectedGroup && groupFolderHint ? <span style={styles.actionHint}>{groupFolderHint}</span> : null}
                    {selectedEmail ? <span style={styles.actionHint}>Email ativo: {selectedEmail.subject || "(sem assunto)"}</span> : null}
                </div>
                <div style={styles.hintText}>
                    A pasta base e as regras documentais dos grupos ficam em Settings &gt; Grupos.
                </div>
                {!selectedGroup ? <PanelState compact tone="info" title="Sem grupo ativo" description="A secao inferior vai mostrar anexos guardados e anexos disponiveis do email aberto." /> : null}
                {documentsError ? <div style={styles.errorText}>{documentsError}</div> : null}

                {selectedGroup ? (
                    <div style={styles.documentSectionStack}>
                        <div style={styles.documentSubsection}>
                            <div style={styles.documentSubTitle}>Anexos do email aberto</div>
                            <div style={styles.hintText}>
                                Nesta fase, os documentos são guardados a partir do email que tens aberto no add-in.
                            </div>
                            <div style={styles.scrollPaneCandidates}>
                                {!currentAttachmentCandidates.length ? (
                                    <PanelState compact tone="info" title="Sem anexos importáveis" description="Abre um email com anexos úteis para os guardares neste grupo." />
                                ) : null}
                                {currentAttachmentCandidates.map((attachment) => (
                                    <div key={attachment.id} style={styles.documentRow}>
                                        <div style={styles.documentMain}>
                                            <span style={styles.documentIcon}><Icons.Paperclip size={12} /></span>
                                            <div style={styles.documentCopy}>
                                                <div style={styles.documentName}>{attachment.name}</div>
                                                <div style={styles.documentMeta}>
                                                    <span>{attachment.contentType || "Anexo"}</span>
                                                    {formatBytes(attachment.size) ? <span>{formatBytes(attachment.size)}</span> : null}
                                                    <span>{ctx.subject || "(sem assunto)"}</span>
                                                </div>
                                            </div>
                                        </div>
                                        <div style={styles.emailActions}>
                                            <IconButton
                                                title="Guardar no grupo"
                                                icon={<Icons.Save size={12} />}
                                                onClick={() => void handleSaveAttachmentsToGroup([attachment])}
                                                disabled={busyAction}
                                                tone="primary"
                                            />
                                        </div>
                                    </div>
                                ))}
                            </div>
                        </div>

                        <div style={styles.documentSubsection}>
                            <div style={styles.documentSubTitle}>Documentos guardados</div>
                            <div style={styles.scrollPaneBottom}>
                                {documentsLoading && !groupDocuments.length ? (
                                    <PanelState compact tone="info" title="A carregar documentos" description="A listar os documentos já guardados neste grupo." />
                                ) : null}
                                {!documentsLoading && !groupDocuments.length ? (
                                    <PanelState compact tone="info" title="Sem documentos guardados" description="Guarda anexos do email aberto para começares a construir a pasta documental do grupo." />
                                ) : null}
                                {groupDocuments.map((document) => {
                                    const active = makeDocumentKey(document) === makeDocumentKey(selectedDocument || {});
                                    const canAttach = Boolean(document.contentBase64);
                                    return (
                                        <div key={makeDocumentKey(document)} style={active ? styles.documentRowActive : styles.documentRow}>
                                            <button type="button" style={styles.emailSelectArea} onClick={() => setSelectedDocumentId(makeDocumentKey(document))}>
                                                <div style={styles.documentName}>{document.name}</div>
                                                <div style={styles.documentMeta}>
                                                    <span>{document.contentType || "Documento"}</span>
                                                    {formatBytes(document.size) ? <span>{formatBytes(document.size)}</span> : null}
                                                    {document.sourceEmailSubject ? <span>{document.sourceEmailSubject}</span> : null}
                                                </div>
                                            </button>
                                            <div style={styles.emailActions}>
                                                <IconButton title="Download" icon={<Icons.Download size={12} />} onClick={() => handleDownloadDocument(document)} disabled={!document.contentBase64} />
                                                <IconButton title="Anexar ao email em edicao" icon={<Icons.Upload size={12} />} onClick={() => void handleAttachDocument(document)} disabled={!canAttach} />
                                                <IconButton title="Remover documento" icon={<Icons.Trash size={12} />} onClick={() => void handleDeleteDocument(document)} disabled={busyAction} tone="danger" />
                                            </div>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                        {selectedDocument ? (
                            <div style={styles.documentSubsection}>
                                <div style={styles.documentSubTitle}>Preview</div>
                                {selectedDocumentPreview?.kind === "image" ? (
                                    <div style={styles.previewFrame}>
                                        <img src={selectedDocumentPreview.dataUrl} alt={selectedDocument.name} style={styles.previewImage} />
                                    </div>
                                ) : null}
                                {selectedDocumentPreview?.kind === "pdf" ? (
                                    <div style={styles.previewFrame}>
                                        <iframe title={selectedDocument.name} src={selectedDocumentPreview.dataUrl} style={styles.previewIframe} />
                                    </div>
                                ) : null}
                                {selectedDocumentPreview?.kind === "text" ? (
                                    <pre style={styles.previewText}>{selectedDocumentPreview.text}</pre>
                                ) : null}
                                {!selectedDocumentPreview || selectedDocumentPreview.kind === "unsupported" ? (
                                    <PanelState compact tone="info" title="Preview não disponível" description="Este tipo de documento pode ser descarregado ou anexado, mas não tem preview interno nesta fase." />
                                ) : null}
                            </div>
                        ) : null}
                    </div>
                ) : null}
            </Section>
        </div>
    );
};

const panelBorder = "1px solid #DFE1E6";

const styles: Record<string, React.CSSProperties> = {
    root: {
        display: "grid",
        gap: "10px",
        alignContent: "start",
    },
    section: {
        border: panelBorder,
        borderRadius: "10px",
        background: "#FFFFFF",
        display: "grid",
        gap: "8px",
        padding: "10px",
        minWidth: 0,
    },
    sectionHeader: {
        display: "flex",
        alignItems: "flex-start",
        justifyContent: "space-between",
        gap: "8px",
        minWidth: 0,
    },
    sectionTitleWrap: {
        display: "grid",
        gap: "2px",
        minWidth: 0,
    },
    sectionTitle: {
        fontSize: "12px",
        fontWeight: 800,
        color: "#172B4D",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    sectionSubtitle: {
        fontSize: "11px",
        color: "#6B778C",
        lineHeight: 1.4,
        wordBreak: "break-word",
    },
    sectionActions: {
        display: "inline-flex",
        alignItems: "center",
        gap: "6px",
        flexShrink: 0,
    },
    sectionBody: {
        display: "grid",
        gap: "8px",
        minWidth: 0,
    },
    inputStack: {
        display: "grid",
        gap: "6px",
    },
    compactField: {
        display: "grid",
    },
    compactFieldRow: {
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "center",
    },
    input: {
        width: "100%",
        border: panelBorder,
        borderRadius: "8px",
        padding: "8px 10px",
        fontSize: "12px",
        background: "#FAFBFC",
        color: "#172B4D",
        minWidth: 0,
    },
    iconBtn: {
        border: panelBorder,
        background: "#F7F8FA",
        color: "#42526E",
        width: "30px",
        height: "30px",
        borderRadius: "8px",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        flexShrink: 0,
    },
    iconBtnPrimary: {
        border: "1px solid #0747A6",
        background: "#0747A6",
        color: "#FFFFFF",
        width: "30px",
        height: "30px",
        borderRadius: "8px",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        flexShrink: 0,
    },
    iconBtnDanger: {
        border: "1px solid #DE350B",
        background: "#FFF0EB",
        color: "#DE350B",
        width: "30px",
        height: "30px",
        borderRadius: "8px",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        flexShrink: 0,
    },
    scrollPaneTop: {
        maxHeight: "180px",
        overflowY: "auto",
        display: "grid",
        gap: "6px",
        paddingRight: "2px",
    },
    scrollPaneMiddle: {
        maxHeight: "320px",
        overflowY: "auto",
        display: "grid",
        gap: "6px",
        paddingRight: "2px",
    },
    scrollPaneBottom: {
        maxHeight: "220px",
        overflowY: "auto",
        display: "grid",
        gap: "6px",
        paddingRight: "2px",
    },
    scrollPaneCandidates: {
        maxHeight: "160px",
        overflowY: "auto",
        display: "grid",
        gap: "6px",
        paddingRight: "2px",
    },
    documentSectionStack: {
        display: "grid",
        gap: "10px",
    },
    documentSubsection: {
        display: "grid",
        gap: "6px",
    },
    documentSubTitle: {
        fontSize: "11px",
        fontWeight: 800,
        textTransform: "uppercase",
        letterSpacing: "0.04em",
        color: "#42526E",
    },
    groupRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "8px 10px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "center",
        textAlign: "left",
        cursor: "pointer",
    },
    groupRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "8px 10px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "center",
        textAlign: "left",
        cursor: "pointer",
    },
    groupMain: {
        display: "grid",
        gap: "2px",
        minWidth: 0,
    },
    groupName: {
        fontSize: "12px",
        fontWeight: 700,
        color: "#172B4D",
        wordBreak: "break-word",
    },
    groupDesc: {
        fontSize: "11px",
        color: "#6B778C",
        wordBreak: "break-word",
    },
    groupCount: {
        fontSize: "11px",
        fontWeight: 800,
        color: "#0747A6",
        borderRadius: "999px",
        background: "#FFFFFF",
        padding: "2px 8px",
        minWidth: "28px",
        textAlign: "center",
    },
    emailRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "start",
    },
    emailRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "start",
    },
    emailSelectArea: {
        border: "none",
        background: "transparent",
        padding: 0,
        textAlign: "left",
        display: "grid",
        gap: "4px",
        cursor: "pointer",
        minWidth: 0,
    },
    emailSubject: {
        fontSize: "12px",
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.35,
        wordBreak: "break-word",
    },
    emailMeta: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
        fontSize: "11px",
        color: "#6B778C",
    },
    emailTagRow: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
    },
    metaTag: {
        fontSize: "10px",
        color: "#42526E",
        background: "#FFFFFF",
        borderRadius: "999px",
        padding: "2px 7px",
        border: panelBorder,
    },
    emailActions: {
        display: "inline-flex",
        gap: "6px",
        flexShrink: 0,
    },
    actionStrip: {
        display: "flex",
        flexWrap: "wrap",
        gap: "8px",
    },
    actionHint: {
        fontSize: "11px",
        color: "#6B778C",
        background: "#FAFBFC",
        borderRadius: "999px",
        padding: "4px 8px",
        border: panelBorder,
    },
    documentRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "center",
    },
    documentRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "center",
    },
    documentMain: {
        display: "grid",
        gridTemplateColumns: "auto 1fr",
        gap: "8px",
        alignItems: "start",
        minWidth: 0,
    },
    documentIcon: {
        color: "#0747A6",
        display: "inline-flex",
        marginTop: "1px",
    },
    documentCopy: {
        display: "grid",
        gap: "2px",
        minWidth: 0,
    },
    documentName: {
        fontSize: "12px",
        fontWeight: 700,
        color: "#172B4D",
        wordBreak: "break-word",
    },
    documentMeta: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
        fontSize: "11px",
        color: "#6B778C",
    },
    errorText: {
        fontSize: "11px",
        color: "#DE350B",
    },
    hintText: {
        fontSize: "11px",
        color: "#6B778C",
    },
    previewFrame: {
        border: panelBorder,
        borderRadius: "8px",
        overflow: "hidden",
        background: "#F7F8FA",
        minHeight: "160px",
        maxHeight: "260px",
    },
    previewImage: {
        width: "100%",
        height: "100%",
        objectFit: "contain",
        display: "block",
        maxHeight: "260px",
        background: "#FFFFFF",
    },
    previewIframe: {
        width: "100%",
        height: "260px",
        border: "none",
        background: "#FFFFFF",
    },
    previewText: {
        margin: 0,
        padding: "10px",
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        fontSize: "11px",
        lineHeight: 1.45,
        color: "#172B4D",
        maxHeight: "260px",
        overflow: "auto",
        whiteSpace: "pre-wrap",
        wordBreak: "break-word",
    },
};
