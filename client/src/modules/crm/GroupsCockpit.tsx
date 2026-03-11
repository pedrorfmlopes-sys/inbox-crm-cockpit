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
    updateLinkGroup,
} from "@/api";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { addBase64AttachmentToCompose, openGroupExplorer, openLinkedOutlookEmail } from "@/office";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

type CurrentAttachmentCandidate = {
    id: string;
    name: string;
    contentType?: string;
    content: string;
    size?: number;
    suspectedInline?: boolean;
    sourceLabel?: string;
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

function normalizeMessageKey(value: string | undefined): string {
    return String(value || "").trim().toLowerCase().replace(/[<>\s]/g, "");
}

function normalizeGroupStorageProvider(value: string | undefined): "cloud" | "local" | "onedrive" {
    const normalized = String(value || "").trim().toLowerCase();
    if (normalized === "local" || normalized === "onedrive") return normalized;
    return "cloud";
}

function isLikelyInlineAttachment(name: string | undefined, contentType: string | undefined): boolean {
    const lowerName = String(name || "").trim().toLowerCase();
    const lowerType = String(contentType || "").trim().toLowerCase();
    if (!lowerName && !lowerType) return false;
    if (!lowerType.startsWith("image/")) return false;
    return /^image\d+\./.test(lowerName)
        || lowerName.includes("logo")
        || lowerName.includes("signature")
        || lowerName.includes("assinatura")
        || lowerName.includes("facebook")
        || lowerName.includes("instagram")
        || lowerName.includes("linkedin");
}

function buildEmailHoverText(email: RelatedEmailEntry): string {
    return [
        email.subject ? `Assunto: ${email.subject}` : "",
        email.fromName || email.fromEmail ? `De: ${email.fromName || email.fromEmail}` : "",
        formatDate(email.messageDateIso || email.receivedAtIso) ? `Data: ${formatDate(email.messageDateIso || email.receivedAtIso)}` : "",
        Array.isArray(email.attachments) ? `Anexos: ${email.attachments.length}` : "",
        Array.isArray(email.relatedRecords) ? `Registos Odoo: ${email.relatedRecords.length}` : "",
    ].filter(Boolean).join("\n");
}

function buildDocumentHoverText(document: Partial<GroupDocumentEntry>): string {
    return [
        document.name ? `Documento: ${document.name}` : "",
        document.contentType ? `Tipo: ${document.contentType}` : "",
        formatBytes(document.size) ? `Tamanho: ${formatBytes(document.size)}` : "",
        document.sourceEmailSubject ? `Email: ${document.sourceEmailSubject}` : "",
    ].filter(Boolean).join("\n");
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
    const [emailsExpanded, setEmailsExpanded] = useState(true);
    const [documentsExpanded, setDocumentsExpanded] = useState(true);
    const [documentDetailsExpanded, setDocumentDetailsExpanded] = useState(false);
    const [attachmentSource, setAttachmentSource] = useState<"current" | "selected">("current");

    const selectedGroup = useMemo(
        () => groups.find((group) => group.id === selectedGroupId) || null,
        [groups, selectedGroupId]
    );
    const trimmedQuery = String(query || "").trim();
    const showAllAlphabetically = trimmedQuery === "/" || trimmedQuery === "*";
    const showGroupSuggestions = Boolean(trimmedQuery);
    const matchingGroups = useMemo(() => {
        if (showAllAlphabetically) return groups;
        const q = trimmedQuery.toLowerCase();
        return groups.filter((group) =>
            String(group?.name || "").toLowerCase().includes(q)
            || String(group?.description || "").toLowerCase().includes(q)
        );
    }, [groups, showAllAlphabetically, trimmedQuery]);
    const documentsEnabled = selectedGroup?.documentsEnabled !== false;
    const groupStorageProvider = normalizeGroupStorageProvider(settings?.groupStorage.provider);

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
                const nextGroups = await listLinkGroups("/");
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
    }, [reloadToken]);

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
        return (attachments || [])
            .map((attachment, index) => ({
                id: `${attachment.name}:${index}`,
                name: String(attachment.name || "").trim(),
                contentType: String(attachment.contentType || "").trim(),
                content: String(attachment.content || "").trim(),
                size: estimateBase64Size(attachment.content),
                suspectedInline: isLikelyInlineAttachment(attachment.name, attachment.contentType),
                sourceLabel: String(ctx.subject || "").trim(),
            }))
            .filter((attachment) => attachment.name && attachment.content)
            .sort((a, b) => {
                const inlineDelta = Number(Boolean(a.suspectedInline)) - Number(Boolean(b.suspectedInline));
                if (inlineDelta !== 0) return inlineDelta;
                return a.name.localeCompare(b.name, "pt-PT");
            });
    }, [attachments, ctx.subject]);

    const selectedEmailMatchesCurrent = useMemo(() => {
        if (!selectedEmail) return false;
        const currentItemId = String(ctx.itemId || "").trim();
        const selectedItemId = String(selectedEmail.itemId || "").trim();
        if (currentItemId && selectedItemId && currentItemId === selectedItemId) return true;
        const currentMessageId = normalizeMessageKey(ctx.internetMessageId);
        const selectedMessageId = normalizeMessageKey(selectedEmail.internetMessageId);
        if (currentMessageId && selectedMessageId && currentMessageId === selectedMessageId) return true;
        return false;
    }, [ctx.internetMessageId, ctx.itemId, selectedEmail]);

    const selectedEmailAttachmentCandidates = useMemo<CurrentAttachmentCandidate[]>(() => {
        if (!selectedEmail || !Array.isArray(selectedEmail.attachments) || !selectedEmail.attachments.length) return [];
        if (selectedEmailMatchesCurrent) {
            return currentAttachmentCandidates.map((attachment) => ({
                ...attachment,
                sourceLabel: selectedEmail.subject || attachment.sourceLabel || "",
            }));
        }
        return selectedEmail.attachments
            .map((attachment, index) => ({
                id: `selected:${makeEmailKey(selectedEmail)}:${attachment.name || index}`,
                name: String(attachment.name || "").trim(),
                contentType: String(attachment.contentType || "").trim(),
                content: "",
                size: Number(attachment.size || 0) || undefined,
                suspectedInline: isLikelyInlineAttachment(attachment.name, attachment.contentType),
                sourceLabel: String(selectedEmail.subject || "").trim(),
            }))
            .filter((attachment) => attachment.name)
            .sort((a, b) => {
                const inlineDelta = Number(Boolean(a.suspectedInline)) - Number(Boolean(b.suspectedInline));
                if (inlineDelta !== 0) return inlineDelta;
                return a.name.localeCompare(b.name, "pt-PT");
            });
    }, [currentAttachmentCandidates, selectedEmail, selectedEmailMatchesCurrent]);

    const visibleAttachmentCandidates = attachmentSource === "selected" ? selectedEmailAttachmentCandidates : currentAttachmentCandidates;
    const savableAttachmentCandidates = visibleAttachmentCandidates.filter((attachment) => Boolean(attachment.content));
    const currentEmailIdentity = useMemo(
        () => [String(ctx.itemId || "").trim(), normalizeMessageKey(ctx.internetMessageId), String(ctx.conversationId || "").trim()].join("|"),
        [ctx.conversationId, ctx.internetMessageId, ctx.itemId]
    );
    const attachmentListKey = `${attachmentSource}:${currentEmailIdentity}:${selectedEmailKey}`;

    const groupFolderHint = useMemo(() => {
        const base = String(settings?.groupStorage.baseFolderPath || "").trim();
        const groupName = String(selectedGroup?.name || "").trim();
        if (!groupName || !documentsEnabled) return "";
        if (!base) return sanitizePathSegment(groupName);
        const separator = /^https?:\/\//i.test(base) || base.endsWith("/") ? "/" : base.includes("\\") ? "\\" : "/";
        return `${base.replace(/[\\/]+$/, "")}${separator}${sanitizePathSegment(groupName)}`;
    }, [documentsEnabled, selectedGroup?.name, settings?.groupStorage.baseFolderPath]);

    useEffect(() => {
        setAttachmentSource("current");
    }, [currentEmailIdentity]);

    async function refreshGroupsAndEmails() {
        setReloadToken((value) => value + 1);
    }

    async function handleCreateGroup() {
        const name = trimmedQuery;
        if (!name || showAllAlphabetically) return;
        const existing = groups.find((group) => String(group.name || "").trim().toLowerCase() === name.toLowerCase());
        if (existing) {
            setSelectedGroupId(existing.id);
            setQuery("");
            setMsg(`Grupo "${existing.name}" selecionado.`);
            return;
        }
        setBusyAction(true);
        try {
            const group = await createLinkGroup({ name, documentsEnabled: true });
            setQuery("");
            setSelectedGroupId(group.id);
            setReloadToken((value) => value + 1);
            setMsg(`Grupo "${group.name}" criado.`);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel criar o grupo.");
        } finally {
            setBusyAction(false);
        }
    }

    async function handleToggleGroupDocuments() {
        if (!selectedGroup) return;
        setBusyAction(true);
        try {
            const nextGroup = await updateLinkGroup(selectedGroup.id, {
                documentsEnabled: !documentsEnabled,
            });
            setGroups((current) => current.map((group) => (group.id === nextGroup.id ? { ...group, ...nextGroup } : group)));
            setMsg(
                nextGroup.documentsEnabled === false
                    ? `Gestao documental desativada no grupo "${nextGroup.name}".`
                    : `Gestao documental ativada no grupo "${nextGroup.name}".`
            );
            setReloadToken((value) => value + 1);
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel atualizar a configuracao documental do grupo.");
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
        if (!documentsEnabled) {
            setMsg("Ativa a gestão documental deste grupo antes de guardares anexos.");
            return;
        }
        setBusyAction(true);
        try {
            const storageProvider = groupStorageProvider;
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

    async function handleOpenExplorer(overrides?: { emailKey?: string; documentId?: string }) {
        if (!selectedGroup) return;
        try {
            await openGroupExplorer({
                groupId: selectedGroup.id,
                ...(overrides?.emailKey ? { emailKey: overrides.emailKey } : {}),
                ...(overrides?.documentId ? { documentId: overrides.documentId } : {}),
            });
        } catch (error: any) {
            setMsg(error?.message || "Nao foi possivel abrir o explorador documental.");
        }
    }

    return (
        <div style={styles.root}>
            <Section
                title="Grupos"
                subtitle="Pesquisa grupos existentes ou cria um novo no mesmo campo."
                actions={
                    <>
                        <IconButton
                            title={selectedGroup ? "Ativar ou desativar documentos deste grupo" : "Seleciona um grupo para gerir documentos"}
                            icon={<Icons.Files size={13} />}
                            onClick={selectedGroup ? () => void handleToggleGroupDocuments() : undefined}
                            disabled={!selectedGroup || busyAction}
                            tone={documentsEnabled ? "primary" : "default"}
                        />
                        <IconButton
                            title={selectedGroup ? "Abrir explorador documental deste grupo" : "Seleciona um grupo para abrir o explorador"}
                            icon={<Icons.ExternalLink size={13} />}
                            onClick={selectedGroup ? () => void handleOpenExplorer() : undefined}
                            disabled={!selectedGroup || busyAction}
                        />
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
                    <div style={styles.compactFieldRow}>
                        <input
                            style={styles.input}
                            value={query}
                            onChange={(event) => setQuery(event.target.value)}
                            onKeyDown={(event) => {
                                if (event.key === "Enter") {
                                    event.preventDefault();
                                    void handleCreateGroup();
                                }
                            }}
                            placeholder='Pesquisar grupos... ("/" mostra todos)'
                        />
                        <IconButton
                            title={trimmedQuery && !showAllAlphabetically ? "Criar grupo com este nome" : "Escreve um nome para criar grupo"}
                            icon={<Icons.Plus size={13} />}
                            onClick={handleCreateGroup}
                            disabled={busyAction || !trimmedQuery || showAllAlphabetically}
                            tone="primary"
                        />
                    </div>
                </div>

                {groupsError ? <div style={styles.errorText}>{groupsError}</div> : null}

                {selectedGroup ? (
                    <div
                        style={styles.selectedGroupCard}
                        title={`${selectedGroup.name}\n${documentsEnabled ? "Documentos ativos" : "Documentos desativados"}\n${selectedGroup.memberCount || 0} email(s)`}
                    >
                        <div style={styles.groupMain}>
                            <div style={styles.groupName}>{selectedGroup.name}</div>
                            <div style={styles.groupMutedMeta}>
                                {documentsEnabled ? "Documentos ativos" : "Documentos desativados"} · {selectedGroup.memberCount || 0} email(s)
                            </div>
                        </div>
                        <div style={styles.selectedGroupActions}>
                            <span style={styles.groupCount}>{selectedGroup.memberCount || 0}</span>
                            <IconButton
                                title="Limpar grupo selecionado"
                                icon={<Icons.RotateCcw size={12} />}
                                onClick={() => {
                                    setSelectedGroupId("");
                                    setSelectedEmailKey("");
                                    setSelectedDocumentId("");
                                }}
                                disabled={busyAction}
                            />
                        </div>
                    </div>
                ) : (
                    <div style={styles.hintText}>Escreve para procurar. Usa "/" se quiseres ver todos os grupos por ordem alfabética.</div>
                )}

                {showGroupSuggestions ? (
                    <div style={styles.scrollPaneTop}>
                        {groupsLoading && !matchingGroups.length ? <PanelState compact tone="info" title="A carregar grupos" description="A procurar grupos que correspondem ao texto." /> : null}
                        {!groupsLoading && !matchingGroups.length ? <PanelState compact tone="info" title="Sem resultados" description="Carrega em + para criar um grupo com este nome." /> : null}
                        {matchingGroups.map((group) => {
                            const selected = group.id === selectedGroupId;
                            return (
                                <button
                                    key={group.id}
                                    type="button"
                                    style={selected ? styles.groupRowActive : styles.groupRow}
                                    onClick={() => {
                                        setSelectedGroupId(group.id);
                                        setQuery("");
                                    }}
                                >
                                    <span style={styles.groupMain}>
                                        <span style={styles.groupName}>{group.name}</span>
                                        <span style={styles.groupDesc}>
                                            {group.documentsEnabled === false ? "Sem documentos" : "Documentos ativos"}
                                        </span>
                                    </span>
                                    <span style={styles.groupCount}>{group.memberCount || 0}</span>
                                </button>
                            );
                        })}
                    </div>
                ) : null}
            </Section>

            <Section
                title="Emails"
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
                        <IconButton
                            title={emailsExpanded ? "Recolher emails" : "Expandir emails"}
                            icon={emailsExpanded ? <Icons.ArrowUp size={10} /> : <Icons.ArrowDown size={10} />}
                            onClick={() => setEmailsExpanded((value) => !value)}
                            disabled={!selectedGroup && !groupEmails.length}
                        />
                    </>
                }
            >
                {selectedGroup ? <div style={styles.sectionMetaHint}>Grupo: {selectedGroup.name}</div> : null}
                {emailsError ? <div style={styles.errorText}>{emailsError}</div> : null}
                {emailsExpanded ? <div style={styles.scrollPaneMiddle}>
                    {!selectedGroup ? <PanelState compact tone="info" title="Nenhum grupo selecionado" description="Escolhe um grupo acima para ver ou gerir os emails." /> : null}
                    {selectedGroup && emailsLoading && !groupEmails.length ? <PanelState compact tone="info" title="A carregar emails" description="A listar os emails ligados ao grupo." /> : null}
                    {selectedGroup && !emailsLoading && !groupEmails.length ? <PanelState compact tone="info" title="Grupo sem emails" description="Podes associar o email atualmente aberto com o botao de ligacao." /> : null}
                    {groupEmails.map((email) => {
                        const active = makeEmailKey(email) === makeEmailKey(selectedEmail || {});
                        const attachmentCount = Array.isArray(email.attachments) ? email.attachments.length : 0;
                        const canOpen = Boolean(email.itemId || email.emailWebLink);
                        return (
                            <div key={makeEmailKey(email)} style={active ? styles.emailRowActive : styles.emailRow}>
                                <button
                                    type="button"
                                    style={styles.emailSelectArea}
                                    onClick={() => setSelectedEmailKey(makeEmailKey(email))}
                                    title={buildEmailHoverText(email)}
                                >
                                    <div style={styles.emailSubject}>{email.subject || "(sem assunto)"}</div>
                                    <div style={styles.emailTagRow}>
                                        {attachmentCount ? <span style={styles.metaTag} title={`${attachmentCount} anexo(s)`}>{attachmentCount}</span> : null}
                                        {Array.isArray(email.relatedRecords) && email.relatedRecords.length ? (
                                            <span style={styles.metaTag} title={`${email.relatedRecords.length} registo(s) Odoo`}>{email.relatedRecords.length}</span>
                                        ) : null}
                                    </div>
                                </button>
                                <div style={styles.emailActions}>
                                    <IconButton title={canOpen ? "Abrir email" : "Sem abertura direta disponivel"} icon={<Icons.MessageSquare size={10} />} onClick={canOpen ? () => void handleOpenEmail(email) : undefined} disabled={!canOpen} />
                                    <IconButton title="Abrir explorador neste email" icon={<Icons.ExternalLink size={10} />} onClick={() => void handleOpenExplorer({ emailKey: makeEmailKey(email) })} disabled={busyAction} />
                                    <IconButton title="Remover do grupo" icon={<Icons.Trash size={10} />} onClick={() => void handleRemoveEmail(email)} disabled={busyAction} tone="danger" />
                                </div>
                            </div>
                        );
                    })}
                </div> : <div style={styles.collapsedHint}>Emails recolhidos</div>}
            </Section>

            <Section
                title="Documentos e acoes"
                subtitle={selectedGroup ? "Documentos guardados no grupo e anexos disponiveis do email aberto." : "Seleciona um grupo para começar a gerir documentos."}
                actions={
                    <>
                        <IconButton
                            title={
                                selectedGroup && savableAttachmentCandidates.length
                                    ? attachmentSource === "selected"
                                        ? "Guardar anexos disponiveis do email selecionado no grupo"
                                        : "Guardar todos os anexos do email aberto no grupo"
                                    : attachmentSource === "selected"
                                        ? "O email selecionado so tem metadados dos anexos nesta fase"
                                        : "Abre um email com anexos para os guardar no grupo"
                            }
                            icon={<Icons.Save size={13} />}
                            onClick={selectedGroup && savableAttachmentCandidates.length ? () => void handleSaveAttachmentsToGroup(savableAttachmentCandidates) : undefined}
                            disabled={!selectedGroup || !documentsEnabled || !savableAttachmentCandidates.length || busyAction}
                            tone="primary"
                        />
                        <IconButton
                            title={selectedGroup ? "Atualizar documentos do grupo" : "Seleciona um grupo"}
                            icon={<Icons.RefreshCw size={13} />}
                            onClick={selectedGroup ? refreshGroupsAndEmails : undefined}
                            disabled={!selectedGroup || documentsLoading || busyAction}
                        />
                        <IconButton
                            title={selectedGroup ? "Abrir explorador documental" : "Seleciona um grupo"}
                            icon={<Icons.ExternalLink size={13} />}
                            onClick={selectedGroup ? () => void handleOpenExplorer() : undefined}
                            disabled={!selectedGroup || busyAction}
                        />
                        <IconButton
                            title={documentsExpanded ? "Recolher documentos" : "Expandir documentos"}
                            icon={documentsExpanded ? <Icons.ArrowUp size={10} /> : <Icons.ArrowDown size={10} />}
                            onClick={() => setDocumentsExpanded((value) => !value)}
                            disabled={!selectedGroup && !groupDocuments.length}
                        />
                    </>
                }
            >
                <div style={styles.actionStripCompact}>
                    <span style={styles.actionHint} title={selectedGroup ? `${groupDocuments.length} documento(s) guardado(s)` : "Sem grupo ativo"}>
                        {selectedGroup ? `${groupDocuments.length}` : "-"}
                    </span>
                    {selectedGroup ? <span style={styles.actionHint} title={documentsEnabled ? "Documentos ativos" : "Documentos desativados"}>{documentsEnabled ? "on" : "off"}</span> : null}
                    <button
                        type="button"
                        style={styles.detailToggle}
                        onClick={() => setDocumentDetailsExpanded((value) => !value)}
                        title={documentDetailsExpanded ? "Esconder detalhes" : "Mostrar detalhes"}
                        disabled={!selectedGroup}
                    >
                        {documentDetailsExpanded ? "menos" : "mais"}
                    </button>
                </div>
                {documentDetailsExpanded && selectedGroup ? (
                    <div style={styles.detailChipWrap}>
                        {groupFolderHint ? <span style={styles.actionHint} title={groupFolderHint}>{groupFolderHint}</span> : null}
                        {selectedEmail ? <span style={styles.actionHint} title={selectedEmail.subject || "(sem assunto)"}>Email ativo</span> : null}
                        <span style={styles.actionHint} title={"As regras documentais dos grupos ficam em Settings > Grupos."}>Settings &gt; Grupos</span>
                    </div>
                ) : null}
                {!selectedGroup ? <PanelState compact tone="info" title="Sem grupo ativo" description="A secao inferior vai mostrar anexos do email aberto ou do email selecionado no grupo." /> : null}
                {documentsError ? <div style={styles.errorText}>{documentsError}</div> : null}

                {selectedGroup && documentsExpanded ? (
                    <div style={styles.documentSectionStack}>
                        {!documentsEnabled ? (
                            <PanelState compact tone="info" title="Gestao documental desativada" description="Este grupo pode continuar a organizar emails, mas nao vai guardar ficheiros ate reativares os documentos." />
                        ) : null}
                        <div style={styles.documentSubsection}>
                            <div style={styles.documentHeaderRow}>
                                <div style={styles.documentSubTitle}>Anexos de origem</div>
                                <div style={styles.sourceSwitch}>
                                    <button
                                        type="button"
                                        style={attachmentSource === "current" ? styles.sourceBtnActive : styles.sourceBtn}
                                        onClick={() => setAttachmentSource("current")}
                                        title="Ver anexos do email atualmente aberto no Outlook"
                                    >
                                        Email aberto
                                    </button>
                                    <button
                                        type="button"
                                        style={attachmentSource === "selected" ? styles.sourceBtnActive : styles.sourceBtn}
                                        onClick={() => setAttachmentSource("selected")}
                                        title="Ver anexos do email selecionado na lista acima"
                                    >
                                        Email selecionado
                                    </button>
                                </div>
                            </div>
                            <div style={styles.hintText}>
                                {attachmentSource === "selected"
                                    ? "Os anexos do email selecionado podem ser revistos aqui. Guardar so fica disponivel quando o conteudo do ficheiro estiver acessivel no add-in."
                                    : "Nesta fase, os documentos sao guardados diretamente a partir do email que tens aberto no add-in."}
                            </div>
                            <div key={attachmentListKey} style={styles.scrollPaneCandidates}>
                                {!documentsEnabled ? (
                                    <PanelState compact tone="info" title="Documentos desativados" description="Ativa os documentos do grupo no topo para guardares anexos." />
                                ) : null}
                                {documentsEnabled && !visibleAttachmentCandidates.length ? (
                                    <PanelState compact tone="info" title="Sem anexos disponiveis" description={attachmentSource === "selected" ? "Seleciona acima um email do grupo que tenha anexos." : "Abre um email com anexos para os poderes guardar neste grupo."} />
                                ) : null}
                                {documentsEnabled ? visibleAttachmentCandidates.map((attachment) => (
                                    <div key={attachment.id} style={styles.documentRow}>
                                        <div style={styles.documentMain}>
                                            <span style={styles.documentIcon}><Icons.Paperclip size={12} /></span>
                                            <div
                                                style={styles.documentCopy}
                                                title={[
                                                    attachment.name,
                                                    attachment.contentType || "Anexo",
                                                    formatBytes(attachment.size),
                                                    attachment.suspectedInline ? "Possível anexo inline/assinatura" : "",
                                                    attachment.sourceLabel || ctx.subject || "(sem assunto)",
                                                ].filter(Boolean).join("\n")}
                                            >
                                                <div style={styles.documentName}>{attachment.name}</div>
                                                <div style={styles.documentMiniMeta}>
                                                    {formatBytes(attachment.size) || attachment.contentType || "Anexo"}
                                                    {attachment.suspectedInline ? " · inline?" : ""}
                                                    {!attachment.content ? " · sem conteudo" : ""}
                                                </div>
                                            </div>
                                        </div>
                                        <div style={styles.emailActions}>
                                            <IconButton
                                                title={attachment.content ? "Guardar no grupo" : "Este anexo ainda nao tem conteudo disponivel para guardar"}
                                                icon={<Icons.Save size={10} />}
                                                onClick={attachment.content ? () => void handleSaveAttachmentsToGroup([attachment]) : undefined}
                                                disabled={busyAction || !documentsEnabled || !attachment.content}
                                                tone="primary"
                                            />
                                        </div>
                                    </div>
                                )) : null}
                            </div>
                        </div>

                        <div style={styles.documentSubsection}>
                            <div style={styles.documentSubTitle}>Documentos guardados</div>
                            <div style={styles.scrollPaneBottom}>
                                {documentsLoading && !groupDocuments.length ? (
                                    <PanelState compact tone="info" title="A carregar documentos" description="A listar os documentos já guardados neste grupo." />
                                ) : null}
                                {!documentsLoading && !groupDocuments.length ? (
                                    <PanelState compact tone="info" title="Sem documentos guardados" description="Guarda anexos do email aberto ou dos emails selecionados para comecares a construir a pasta documental do grupo." />
                                ) : null}
                                {groupDocuments.map((document) => {
                                    const active = makeDocumentKey(document) === makeDocumentKey(selectedDocument || {});
                                    const canAttach = Boolean(document.contentBase64);
                                    return (
                                        <div key={makeDocumentKey(document)} style={active ? styles.documentRowActive : styles.documentRow}>
                                            <button
                                                type="button"
                                                style={styles.emailSelectArea}
                                                onClick={() => setSelectedDocumentId(makeDocumentKey(document))}
                                                title={buildDocumentHoverText(document)}
                                            >
                                                <div style={styles.documentName}>{document.name}</div>
                                                <div style={styles.documentMiniMeta}>{formatBytes(document.size) || document.contentType || "Documento"}</div>
                                            </button>
                                            <div style={styles.emailActions}>
                                                <IconButton title="Download" icon={<Icons.Download size={10} />} onClick={() => handleDownloadDocument(document)} disabled={!document.contentBase64} />
                                                <IconButton title="Abrir explorador neste documento" icon={<Icons.ExternalLink size={10} />} onClick={() => void handleOpenExplorer({ documentId: document.id })} disabled={busyAction} />
                                                <IconButton title="Anexar ao email em edicao" icon={<Icons.Upload size={10} />} onClick={() => void handleAttachDocument(document)} disabled={!canAttach} />
                                                <IconButton title="Remover documento" icon={<Icons.Trash size={10} />} onClick={() => void handleDeleteDocument(document)} disabled={busyAction} tone="danger" />
                                            </div>
                                        </div>
                                    );
                                })}
                            </div>
                        </div>

                    </div>
                ) : null}
                {selectedGroup && !documentsExpanded ? <div style={styles.collapsedHint}>Documentos recolhidos</div> : null}
            </Section>
        </div>
    );
};

const panelBorder = "1px solid #DFE1E6";

const styles: Record<string, React.CSSProperties> = {
    root: {
        display: "grid",
        gap: "8px",
        alignContent: "start",
    },
    section: {
        border: panelBorder,
        borderRadius: "10px",
        background: "#FFFFFF",
        display: "grid",
        gap: "6px",
        padding: "8px",
        minWidth: 0,
    },
    sectionHeader: {
        display: "flex",
        alignItems: "flex-start",
        justifyContent: "space-between",
        gap: "6px",
        minWidth: 0,
    },
    sectionTitleWrap: {
        display: "grid",
        gap: "2px",
        minWidth: 0,
    },
    sectionTitle: {
        fontSize: "11px",
        fontWeight: 800,
        color: "#172B4D",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    sectionSubtitle: {
        display: "none",
    },
    sectionActions: {
        display: "inline-flex",
        alignItems: "center",
        gap: "4px",
        flexShrink: 0,
    },
    sectionBody: {
        display: "grid",
        gap: "6px",
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
        padding: "7px 9px",
        fontSize: "11px",
        background: "#FAFBFC",
        color: "#172B4D",
        minWidth: 0,
    },
    iconBtn: {
        border: panelBorder,
        background: "#F7F8FA",
        color: "#42526E",
        width: "22px",
        height: "22px",
        borderRadius: "7px",
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
        width: "22px",
        height: "22px",
        borderRadius: "7px",
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
        width: "22px",
        height: "22px",
        borderRadius: "7px",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        flexShrink: 0,
    },
    scrollPaneTop: {
        maxHeight: "136px",
        overflowY: "auto",
        display: "grid",
        gap: "4px",
        paddingRight: "2px",
    },
    scrollPaneMiddle: {
        maxHeight: "268px",
        overflowY: "auto",
        display: "grid",
        gap: "4px",
        paddingRight: "2px",
    },
    scrollPaneBottom: {
        maxHeight: "210px",
        overflowY: "auto",
        display: "grid",
        gap: "4px",
        paddingRight: "2px",
    },
    scrollPaneCandidates: {
        maxHeight: "122px",
        overflowY: "auto",
        display: "grid",
        gap: "4px",
        paddingRight: "2px",
    },
    selectedGroupCard: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#F7FAFF",
        padding: "6px 8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "8px",
        alignItems: "center",
    },
    selectedGroupActions: {
        display: "inline-flex",
        alignItems: "center",
        gap: "4px",
        flexShrink: 0,
    },
    documentSectionStack: {
        display: "grid",
        gap: "8px",
    },
    documentSubsection: {
        display: "grid",
        gap: "4px",
    },
    documentSubTitle: {
        fontSize: "10px",
        fontWeight: 800,
        textTransform: "uppercase",
        letterSpacing: "0.04em",
        color: "#42526E",
    },
    documentHeaderRow: {
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: "6px",
        flexWrap: "wrap",
    },
    sourceSwitch: {
        display: "inline-flex",
        gap: "4px",
        flexWrap: "wrap",
    },
    sourceBtn: {
        border: panelBorder,
        background: "#FFFFFF",
        color: "#42526E",
        borderRadius: "999px",
        padding: "2px 7px",
        fontSize: "9px",
        fontWeight: 700,
        cursor: "pointer",
    },
    sourceBtnActive: {
        border: "1px solid #0747A6",
        background: "#E9F2FF",
        color: "#0747A6",
        borderRadius: "999px",
        padding: "2px 7px",
        fontSize: "9px",
        fontWeight: 700,
        cursor: "pointer",
    },
    groupRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "6px 8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "center",
        textAlign: "left",
        cursor: "pointer",
    },
    groupRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "6px 8px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "center",
        textAlign: "left",
        cursor: "pointer",
    },
    groupMain: {
        display: "grid",
        gap: "1px",
        minWidth: 0,
    },
    groupName: {
        fontSize: "11px",
        fontWeight: 600,
        color: "#172B4D",
        wordBreak: "break-word",
    },
    groupDesc: {
        display: "none",
    },
    groupMutedMeta: {
        display: "none",
    },
    groupCount: {
        fontSize: "10px",
        fontWeight: 800,
        color: "#0747A6",
        borderRadius: "999px",
        background: "#FFFFFF",
        padding: "2px 6px",
        minWidth: "24px",
        textAlign: "center",
    },
    emailRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "6px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "start",
    },
    emailRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "6px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "start",
    },
    emailSelectArea: {
        border: "none",
        background: "transparent",
        padding: 0,
        textAlign: "left",
        display: "grid",
        gap: "3px",
        cursor: "pointer",
        minWidth: 0,
    },
    emailSubject: {
        fontSize: "10.5px",
        fontWeight: 600,
        color: "#172B4D",
        lineHeight: 1.22,
        wordBreak: "break-word",
    },
    emailMeta: {
        display: "none",
    },
    emailTagRow: {
        display: "flex",
        flexWrap: "wrap",
        gap: "4px",
    },
    metaTag: {
        fontSize: "9px",
        color: "#42526E",
        background: "#FFFFFF",
        borderRadius: "999px",
        padding: "1px 6px",
        border: panelBorder,
    },
    emailActions: {
        display: "inline-flex",
        gap: "4px",
        flexShrink: 0,
    },
    actionStrip: {
        display: "flex",
        flexWrap: "wrap",
        gap: "8px",
    },
    actionStripCompact: {
        display: "flex",
        alignItems: "center",
        gap: "6px",
        flexWrap: "wrap",
    },
    actionHint: {
        fontSize: "9px",
        color: "#6B778C",
        background: "#FAFBFC",
        borderRadius: "999px",
        padding: "2px 6px",
        border: panelBorder,
        maxWidth: "100%",
        overflow: "hidden",
        textOverflow: "ellipsis",
        whiteSpace: "nowrap",
    },
    detailToggle: {
        border: panelBorder,
        borderRadius: "999px",
        background: "#FFFFFF",
        color: "#42526E",
        fontSize: "9px",
        padding: "2px 7px",
        cursor: "pointer",
    },
    detailChipWrap: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
    },
    documentRow: {
        border: panelBorder,
        borderRadius: "8px",
        background: "#FAFBFC",
        padding: "6px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "center",
    },
    documentRowActive: {
        border: "1px solid #0747A6",
        borderRadius: "8px",
        background: "#E9F2FF",
        padding: "6px",
        display: "grid",
        gridTemplateColumns: "1fr auto",
        gap: "6px",
        alignItems: "center",
    },
    documentMain: {
        display: "grid",
        gridTemplateColumns: "auto 1fr",
        gap: "6px",
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
        gap: "1px",
        minWidth: 0,
    },
    documentName: {
        fontSize: "10.5px",
        fontWeight: 600,
        color: "#172B4D",
        wordBreak: "break-word",
    },
    documentMeta: {
        display: "none",
    },
    documentMiniMeta: {
        fontSize: "9px",
        color: "#6B778C",
        lineHeight: 1.2,
    },
    errorText: {
        fontSize: "11px",
        color: "#DE350B",
    },
    hintText: {
        display: "none",
    },
    mutedHint: {
        fontSize: "10px",
        color: "#6B778C",
    },
    sectionMetaHint: {
        fontSize: "10px",
        color: "#6B778C",
    },
    collapsedHint: {
        fontSize: "10px",
        color: "#6B778C",
        padding: "2px 0",
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
