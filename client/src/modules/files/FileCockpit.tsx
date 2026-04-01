import React, { useEffect, useMemo, useRef, useState } from "react";
import * as Icons from "../../ui/icons";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { getEmailAttachmentContentBase64, getInvoiceStudioBatchStatus, getRelatedEmailContext, uploadToInvoiceStudio } from "@/api";
import { PanelState, type PanelStateTone } from "../../ui/PanelState";

type InvoiceJobState = {
    batchId: string;
    status: string;
    total: number;
    done: number;
    errors: number;
    rows: Array<Record<string, any>>;
};

function isPdfLike(file: { name?: string; type?: string }): boolean {
    const mime = String(file?.type || "").toLowerCase();
    const name = String(file?.name || "").toLowerCase();
    return mime.includes("pdf") || name.endsWith(".pdf");
}

export const FileCockpit: React.FC = () => {
    const {
        files,
        addFile,
        removeFile,
        setMsg,
        setTab,
        setAiState,
        settings,
        openSettingsSection,
        ctx,
    } = useCockpit() as any;

    const [isDragging, setIsDragging] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [isSendingInvoiceStudio, setIsSendingInvoiceStudio] = useState(false);
    const [selectedForInvoiceStudio, setSelectedForInvoiceStudio] = useState<Record<string, boolean>>({});
    const [invoiceStudioJob, setInvoiceStudioJob] = useState<InvoiceJobState | null>(null);
    const [status, setStatus] = useState<{ tone: PanelStateTone; title: string; description?: string } | null>(null);
    const hiddenFileInput = useRef<HTMLInputElement>(null);

    const invoiceStudioEnabled = settings?.invoiceStudio?.enabled === true;
    const invoiceStudioMissingFields = useMemo(() => {
        const missing: string[] = [];
        if (!String(settings?.invoiceStudio?.baseUrl || "").trim()) missing.push("URL");
        if (!String(settings?.invoiceStudio?.email || "").trim()) missing.push("email");
        if (!String(settings?.invoiceStudio?.password || "").trim()) missing.push("password");
        return missing;
    }, [settings?.invoiceStudio?.baseUrl, settings?.invoiceStudio?.email, settings?.invoiceStudio?.password]);
    const invoiceStudioReady = invoiceStudioEnabled && invoiceStudioMissingFields.length === 0;

    useEffect(() => {
        setSelectedForInvoiceStudio((prev) => {
            const next: Record<string, boolean> = {};
            for (const file of files || []) {
                const name = String(file?.name || "").trim();
                if (!name) continue;
                next[name] = typeof prev[name] === "boolean" ? prev[name] : true;
            }
            return next;
        });
    }, [files]);

    const selectedInvoiceFiles = useMemo(
        () => (files || []).filter((file: any) => selectedForInvoiceStudio[String(file?.name || "").trim()] !== false),
        [files, selectedForInvoiceStudio]
    );

    const selectedInvoicePdfFiles = useMemo(
        () => selectedInvoiceFiles.filter((file: any) => isPdfLike(file)),
        [selectedInvoiceFiles]
    );

    const handleDragOver = (e: React.DragEvent) => {
        e.preventDefault();
        setIsDragging(true);
    };

    const handleDragLeave = () => {
        setIsDragging(false);
    };

    const processFiles = async (fileList: File[]) => {
        for (const file of fileList) {
            const reader = new FileReader();
            reader.onload = () => {
                const base64 = (reader.result as string).split(",")[1];
                addFile({
                    name: file.name,
                    type: file.type,
                    content: base64,
                });
            };
            reader.readAsDataURL(file);
        }
        return fileList.length;
    };

    const handleDrop = async (e: React.DragEvent) => {
        e.preventDefault();
        setIsDragging(false);
        const droppedFiles = Array.from(e.dataTransfer.files);
        const added = await processFiles(droppedFiles);
        if (added > 0) {
            setStatus({
                tone: "success",
                title: "Ficheiros adicionados",
                description: `${added} ficheiro(s) prontos para analise pela IA ou envio para o InvoiceStudio.`,
            });
        }
    };

    const onFileChange = async (e: React.ChangeEvent<HTMLInputElement>) => {
        const selectedFiles = Array.from(e.target.files || []);
        const added = await processFiles(selectedFiles);
        if (added > 0) {
            setStatus({
                tone: "success",
                title: "Ficheiros carregados",
                description: `${added} ficheiro(s) adicionados a esta conversa.`,
            });
        }
        e.target.value = "";
    };

    const triggerFileSummary = (fileName: string) => {
        setAiState({ prompt: `ANALISE DETALHADA DO ANEXO "${fileName}": extrai datas, valores, numero de faturas e descreve o conteudo tecnico.` });
        setTab("ai");
    };

    const handleImportAttachments = async () => {
        try {
            setIsImporting(true);
            setStatus({
                tone: "loading",
                title: "A importar anexos",
                description: "Estamos a ler os anexos persistidos do email atual no servidor.",
            });
            let imported = 0;
            const persisted = await getRelatedEmailContext({
                itemId: String(ctx?.itemId || "").trim(),
                internetMessageId: String(ctx?.internetMessageId || "").trim(),
                conversationId: String(ctx?.conversationId || "").trim(),
                subject: String(ctx?.subject || "").trim(),
                fromEmail: String(ctx?.fromEmail || "").trim(),
                receivedAtIso: String(ctx?.receivedDateTimeIso || "").trim(),
            }).catch(() => null);

            const persistedEmail = persisted?.email || null;
            const persistedAttachments = Array.isArray(persistedEmail?.attachments) ? persistedEmail.attachments : [];
            if (persistedEmail?.id && persistedAttachments.length) {
                for (const attachment of persistedAttachments) {
                    let content = String(attachment?.content || "").trim();
                    if (!content && attachment?.hasContent) {
                        const remoteId = String(attachment?.key || attachment?.id || "").trim();
                        if (remoteId) {
                            try {
                                const loaded = await getEmailAttachmentContentBase64(String(persistedEmail.id || "").trim(), remoteId);
                                content = String(loaded.base64 || "").trim();
                            } catch {
                                content = "";
                            }
                        }
                    }
                    if (!content) continue;
                    addFile({
                        name: String(attachment?.name || "").trim(),
                        type: String(attachment?.contentType || "application/octet-stream").trim(),
                        content,
                    });
                    imported += 1;
                }
            }

            if (imported === 0) {
                const emailWasPersisted = Boolean(persistedEmail?.id);
                if (emailWasPersisted) {
                    setMsg("Nenhum anexo persistido encontrado neste email.");
                    setStatus({
                        tone: "empty",
                        title: "Sem anexos persistidos",
                        description: "Este email ainda nao tem anexos disponiveis na persistencia central, ou nao tem anexos para importar.",
                    });
                } else {
                    setMsg("O email atual ainda nao esta pronto no servidor.");
                    setStatus({
                        tone: "error",
                        title: "Email ainda nao persistido",
                        description: "Volta a abrir este email daqui a pouco ou reabre a app para concluir a ingestao antes de importar anexos.",
                    });
                }
            } else {
                setMsg(`${imported} anexos importados com sucesso.`);
                setStatus({
                    tone: "success",
                    title: "Anexos importados",
                    description: `${imported} anexo(s) disponiveis para leitura, resumo ou envio para o InvoiceStudio.`,
                });
            }
        } catch (e: any) {
            console.error("Import error", e);
            setMsg("Erro ao importar anexos: " + e.message);
            setStatus({
                tone: "error",
                title: "Falha ao importar anexos",
                description: e?.message || "Nao foi possivel ler os anexos deste email.",
            });
        } finally {
            setIsImporting(false);
        }
    };

    const toggleInvoiceSelection = (fileName: string) => {
        setSelectedForInvoiceStudio((prev) => ({
            ...prev,
            [fileName]: !prev[fileName],
        }));
    };

    const selectAllInvoiceFiles = (checked: boolean) => {
        const next: Record<string, boolean> = {};
        for (const file of files || []) {
            const name = String(file?.name || "").trim();
            if (!name) continue;
            next[name] = checked;
        }
        setSelectedForInvoiceStudio(next);
    };

    const buildInvoiceBatchId = () => {
        const seed = String(ctx?.internetMessageId || ctx?.itemId || ctx?.subject || "email")
            .trim()
            .toLowerCase()
            .replace(/[^a-z0-9]+/g, "-")
            .replace(/^-+|-+$/g, "")
            .slice(0, 48) || "email";
        return `icc-${seed}-${Date.now()}`;
    };

    const pollInvoiceStudioBatch = async (batchId: string) => {
        const credentials = {
            baseUrl: String(settings?.invoiceStudio?.baseUrl || "").trim(),
            email: String(settings?.invoiceStudio?.email || "").trim(),
            password: String(settings?.invoiceStudio?.password || "").trim(),
            project: String(settings?.invoiceStudio?.project || "").trim(),
            batchId,
        };

        for (let attempt = 0; attempt < 25; attempt += 1) {
            if (attempt > 0) {
                await new Promise((resolve) => setTimeout(resolve, 3000));
            }

            const snapshot = await getInvoiceStudioBatchStatus(credentials);
            const progress = snapshot.progress || {};
            const statusLabel = String(progress.status || "").trim() || "processing";
            setInvoiceStudioJob({
                batchId,
                status: statusLabel,
                total: Number(progress.total || 0),
                done: Number(progress.done || 0),
                errors: Number(progress.errors || 0),
                rows: Array.isArray(snapshot.rows) ? snapshot.rows : [],
            });

            if (statusLabel.toLowerCase() === "finished" || (Number(progress.total || 0) > 0 && Number(progress.done || 0) >= Number(progress.total || 0))) {
                const rowCount = Array.isArray(snapshot.rows) ? snapshot.rows.length : 0;
                setStatus({
                    tone: Number(progress.errors || 0) > 0 ? "info" : "success",
                    title: Number(progress.errors || 0) > 0 ? "Processamento concluido com avisos" : "InvoiceStudio concluido",
                    description: rowCount
                        ? `${rowCount} documento(s) disponiveis no batch ${batchId}.`
                        : `Batch ${batchId} concluido no InvoiceStudio.`,
                });
                return;
            }
        }

        setStatus({
            tone: "info",
            title: "Processamento em curso",
            description: `O batch ${batchId} continua em processamento. Podes voltar a esta aba para consultar o estado.`,
        });
    };

    const handleSendToInvoiceStudio = async () => {
        if (!invoiceStudioEnabled) {
            setStatus({
                tone: "info",
                title: "InvoiceStudio desligado",
                description: "Ativa esta integracao em Settings > Ligacoes.",
            });
            setMsg("Ativa o modulo InvoiceStudio em Settings > Ligacoes.");
            openSettingsSection("conns");
            return;
        }

        if (!invoiceStudioReady) {
            setStatus({
                tone: "warning",
                title: "Configuracao incompleta",
                description: `Falta configurar: ${invoiceStudioMissingFields.join(", ")}. Abre Settings > Ligacoes para completar.`,
            });
            setMsg(`InvoiceStudio: falta configurar ${invoiceStudioMissingFields.join(", ")}.`);
            openSettingsSection("conns");
            return;
        }

        if (!selectedInvoicePdfFiles.length) {
            setStatus({
                tone: "warning",
                title: "Sem PDFs selecionados",
                description: "O InvoiceStudio atual processa sobretudo PDFs. Seleciona pelo menos um PDF.",
            });
            return;
        }

        const batchId = buildInvoiceBatchId();
        setIsSendingInvoiceStudio(true);
        setStatus({
            tone: "loading",
            title: "A enviar para o InvoiceStudio",
            description: `A preparar ${selectedInvoicePdfFiles.length} PDF(s) para processamento.`,
        });

        try {
            const response = await uploadToInvoiceStudio({
                baseUrl: String(settings?.invoiceStudio?.baseUrl || "").trim(),
                email: String(settings?.invoiceStudio?.email || "").trim(),
                password: String(settings?.invoiceStudio?.password || "").trim(),
                project: String(settings?.invoiceStudio?.project || "").trim(),
                batchId,
                metadata: {
                    subject: String(ctx?.subject || "").trim(),
                    fromEmail: String(ctx?.fromEmail || "").trim(),
                    fromName: String(ctx?.fromName || "").trim(),
                    conversationId: String(ctx?.conversationId || "").trim(),
                    internetMessageId: String(ctx?.internetMessageId || "").trim(),
                    itemId: String(ctx?.itemId || "").trim(),
                    receivedAtIso: String(ctx?.receivedDateTimeIso || "").trim(),
                },
                files: selectedInvoicePdfFiles.map((file: any) => ({
                    name: String(file?.name || "").trim(),
                    type: String(file?.type || "").trim() || "application/pdf",
                    content: String(file?.content || "").trim(),
                })),
            });

            setInvoiceStudioJob({
                batchId: response.batchId,
                status: String(response.status || "processing"),
                total: Number(response.count || selectedInvoicePdfFiles.length),
                done: 0,
                errors: 0,
                rows: [],
            });

            setStatus({
                tone: "loading",
                title: "Batch criado",
                description: `Batch ${response.batchId} enviado. A acompanhar o processamento no InvoiceStudio.`,
            });

            await pollInvoiceStudioBatch(response.batchId);
        } catch (error: any) {
            console.error("[Files] InvoiceStudio upload failed", error);
            setStatus({
                tone: "error",
                title: "Falha no envio",
                description: error?.message || "Nao foi possivel enviar os ficheiros para o InvoiceStudio.",
            });
            setMsg(`InvoiceStudio: ${error?.message || "falha no envio"}`);
        } finally {
            setIsSendingInvoiceStudio(false);
        }
    };

    return (
        <div style={S.container}>
            <div style={S.header}>
                <h3 style={S.title}>Documentos e Anexos</h3>
                <p style={S.subtitle}>Importa anexos do email, analisa-os com a IA ou envia-os para o InvoiceStudio.</p>
            </div>

            {status && (
                <PanelState
                    tone={status.tone}
                    title={status.title}
                    description={status.description}
                    compact
                />
            )}

            <div style={{ display: "flex", gap: "10px" }}>
                <button
                    style={S.importBtn}
                    onClick={handleImportAttachments}
                    disabled={isImporting}
                    title="Importar do Email"
                >
                    {isImporting ? <Icons.RotateCcw size={16} style={{ animation: "spin 1s linear infinite" }} /> : <Icons.Link size={16} />}
                    <span>{isImporting ? "A importar..." : "Importar Anexos"}</span>
                </button>

                <button
                    style={S.browseBtn}
                    onClick={() => hiddenFileInput.current?.click()}
                    title="Carregar do PC"
                >
                    <Icons.Plus size={16} />
                    <span>Carregar PC</span>
                </button>
            </div>

            <div style={S.invoiceCard}>
                <div style={S.invoiceHeader}>
                    <div>
                        <div style={S.fieldLabel}>InvoiceStudio</div>
                        <div style={S.fieldHint}>Modulo isolado e opcional para enviar PDFs desta aba para processamento.</div>
                    </div>
                    <button style={S.smallGhostBtn} type="button" onClick={() => openSettingsSection("conns")}>
                        Ligacoes
                    </button>
                </div>

                {!invoiceStudioEnabled && (
                    <PanelState
                        tone="info"
                        title="Integracao desativada"
                        description="Ativa o InvoiceStudio em Settings > Ligacoes para usar este envio."
                        compact
                    />
                )}

                {invoiceStudioEnabled && !invoiceStudioReady && (
                    <PanelState
                        tone="warning"
                        title="Configuracao incompleta"
                        description={`Falta configurar: ${invoiceStudioMissingFields.join(", ")}.`}
                        compact
                    />
                )}

                {invoiceStudioEnabled && (
                    <div style={{ display: "flex", alignItems: "center", justifyContent: "space-between", gap: 8 }}>
                        <div style={{ fontSize: 11, color: "var(--iccc-text-muted)" }}>
                            {selectedInvoicePdfFiles.length} PDF(s) selecionado(s) para envio
                        </div>
                        <div style={{ display: "flex", gap: 8 }}>
                            <button style={S.smallGhostBtn} type="button" onClick={() => selectAllInvoiceFiles(true)}>Todos</button>
                            <button style={S.smallGhostBtn} type="button" onClick={() => selectAllInvoiceFiles(false)}>Nenhum</button>
                        </div>
                    </div>
                )}

                {invoiceStudioEnabled && (
                    <button
                        style={S.invoiceBtn}
                        type="button"
                        onClick={handleSendToInvoiceStudio}
                        disabled={isSendingInvoiceStudio || selectedInvoicePdfFiles.length === 0}
                    >
                        <Icons.ExternalLink size={16} />
                        <span>
                            {isSendingInvoiceStudio
                                ? "A enviar..."
                                : !invoiceStudioEnabled
                                    ? "Ativar InvoiceStudio"
                                    : !invoiceStudioReady
                                        ? "Configurar InvoiceStudio"
                                        : "Enviar para InvoiceStudio"}
                        </span>
                    </button>
                )}

                {invoiceStudioJob && (
                    <div style={S.invoiceJobCard}>
                        <div style={S.statsRow}>
                            <span>Batch</span>
                            <span style={S.statsVal}>{invoiceStudioJob.batchId}</span>
                        </div>
                        <div style={S.statsRow}>
                            <span>Estado</span>
                            <span style={S.statsVal}>{invoiceStudioJob.status}</span>
                        </div>
                        <div style={S.statsRow}>
                            <span>Progresso</span>
                            <span style={S.statsVal}>{invoiceStudioJob.done}/{invoiceStudioJob.total || "?"}</span>
                        </div>
                        {invoiceStudioJob.errors > 0 && (
                            <div style={S.statsRow}>
                                <span>Erros</span>
                                <span style={{ ...S.statsVal, color: "#ef4444" }}>{invoiceStudioJob.errors}</span>
                            </div>
                        )}
                    </div>
                )}
            </div>

            <div
                style={{
                    ...S.dropZone,
                    borderColor: isDragging ? "var(--iccc-pill-active-bg)" : "var(--iccc-card-border)",
                    background: isDragging ? "rgba(59, 130, 246, 0.1)" : "var(--iccc-card-bg)",
                }}
                onDragOver={handleDragOver}
                onDragLeave={handleDragLeave}
                onDrop={handleDrop}
            >
                <input
                    type="file"
                    ref={hiddenFileInput}
                    style={{ display: "none" }}
                    onChange={onFileChange}
                    multiple
                />
                <div style={{ pointerEvents: "none", display: "flex", flexDirection: "column", alignItems: "center", gap: "8px" }}>
                    <Icons.Files size={24} color="var(--iccc-text-muted)" />
                    <span style={{ fontSize: "12px", color: "var(--iccc-text-muted)" }}>Arrasta ficheiros para aqui</span>
                </div>
            </div>

            <div style={S.fileList}>
                {files.length === 0 && (
                    <PanelState
                        tone="empty"
                        title="Nenhum ficheiro carregado"
                        description="Importa anexos do email ou adiciona ficheiros do computador para os analisar."
                    />
                )}

                {files.map((file: any, idx: number) => {
                    const fileName = String(file?.name || "").trim();
                    const checked = selectedForInvoiceStudio[fileName] !== false;
                    return (
                        <div key={idx} style={S.fileItem}>
                            <label style={S.checkboxWrap} title="Selecionar para envio ao InvoiceStudio">
                                <input
                                    type="checkbox"
                                    checked={checked}
                                    onChange={() => toggleInvoiceSelection(fileName)}
                                />
                            </label>
                            <div style={S.fileIcon}>
                                {isPdfLike(file) ? <Icons.Receipt size={16} color="#ef4444" /> : <Icons.Files size={16} color="#3b82f6" />}
                            </div>
                            <div style={S.fileInfo}>
                                <div style={S.fileName} title={fileName}>{fileName}</div>
                                <div style={S.fileMeta}>
                                    {String(file?.type || "unknown")}
                                    {isPdfLike(file) ? " • PDF" : " • fora do perfil InvoiceStudio"}
                                </div>
                            </div>
                            <button
                                style={S.actionIconBtn}
                                onClick={() => triggerFileSummary(fileName)}
                                title="Resumir este ficheiro na IA"
                            >
                                <Icons.Sparkles size={14} color="var(--iccc-pill-active-bg)" />
                            </button>
                            <button style={S.deleteBtn} onClick={() => removeFile(fileName)} title="Remover">
                                x
                            </button>
                        </div>
                    );
                })}
            </div>

            <div style={S.statsCard}>
                <div style={S.statsRow}>
                    <span>Ficheiros em memoria</span>
                    <span style={S.statsVal}>{files.length}</span>
                </div>
                <div style={S.statsRow}>
                    <span>Selecionados p/ InvoiceStudio</span>
                    <span style={S.statsVal}>{selectedInvoicePdfFiles.length}</span>
                </div>
            </div>
            <style>{`
                @keyframes spin { 100% { transform: rotate(360deg); } }
            `}</style>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    container: {
        display: "flex",
        flexDirection: "column",
        gap: "16px",
        paddingTop: "4px",
    },
    header: {
        textAlign: "center",
        marginBottom: "4px",
    },
    title: {
        fontSize: "14px",
        fontWeight: 800,
        textTransform: "uppercase",
        letterSpacing: "0.05em",
        color: "var(--iccc-text)",
        margin: "0 0 4px 0",
    },
    subtitle: {
        fontSize: "11px",
        color: "var(--iccc-text-muted)",
    },
    fieldLabel: {
        fontSize: "11px",
        fontWeight: 800,
        color: "var(--iccc-text)",
        textTransform: "uppercase",
        letterSpacing: "0.04em",
    },
    fieldHint: {
        fontSize: "11px",
        color: "var(--iccc-text-muted)",
        marginTop: "2px",
    },
    importBtn: {
        flex: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "8px",
        background: "var(--iccc-btn-bg)",
        color: "var(--iccc-btn-text)",
        border: "none",
        borderRadius: "8px",
        padding: "10px",
        fontSize: "12px",
        fontWeight: 600,
        cursor: "pointer",
        transition: "all 0.2s",
        boxShadow: "var(--iccc-shadow)",
    },
    browseBtn: {
        flex: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "8px",
        background: "var(--iccc-bg)",
        color: "var(--iccc-text)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "8px",
        padding: "10px",
        fontSize: "12px",
        fontWeight: 600,
        cursor: "pointer",
        transition: "all 0.2s",
    },
    invoiceCard: {
        display: "grid",
        gap: "10px",
        padding: "12px",
        borderRadius: "12px",
        border: "1px solid var(--iccc-card-border)",
        background: "var(--iccc-card-bg)",
        boxShadow: "var(--iccc-shadow)",
    },
    invoiceHeader: {
        display: "flex",
        alignItems: "flex-start",
        justifyContent: "space-between",
        gap: "10px",
    },
    smallGhostBtn: {
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "8px",
        padding: "6px 10px",
        background: "var(--iccc-bg)",
        color: "var(--iccc-text)",
        fontSize: "11px",
        fontWeight: 700,
        cursor: "pointer",
    },
    invoiceBtn: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "8px",
        background: "#0f766e",
        color: "#fff",
        border: "none",
        borderRadius: "8px",
        padding: "10px 12px",
        fontSize: "12px",
        fontWeight: 700,
        cursor: "pointer",
    },
    invoiceJobCard: {
        background: "rgba(15,118,110,0.06)",
        border: "1px solid rgba(15,118,110,0.18)",
        borderRadius: "10px",
        padding: "10px 12px",
        display: "grid",
        gap: "6px",
    },
    dropZone: {
        border: "2px dashed var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        padding: "24px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        transition: "all 0.2s ease",
        cursor: "default",
        minHeight: "100px",
    },
    fileList: {
        display: "flex",
        flexDirection: "column",
        gap: "8px",
        maxHeight: "240px",
        overflowY: "auto",
    },
    fileItem: {
        display: "flex",
        alignItems: "center",
        padding: "8px",
        background: "var(--iccc-card-bg)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "8px",
        gap: "10px",
    },
    checkboxWrap: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        marginRight: "2px",
    },
    fileIcon: {
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        width: "24px",
        height: "24px",
        background: "rgba(0,0,0,0.03)",
        borderRadius: "4px",
    },
    fileInfo: {
        flex: 1,
        overflow: "hidden",
    },
    fileName: {
        fontSize: "12px",
        fontWeight: 600,
        color: "var(--iccc-text)",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
    },
    fileMeta: {
        fontSize: "10px",
        color: "var(--iccc-text-muted)",
    },
    deleteBtn: {
        background: "none",
        border: "none",
        color: "var(--iccc-text-muted)",
        fontSize: "14px",
        cursor: "pointer",
        padding: "4px",
    },
    actionIconBtn: {
        background: "rgba(59, 130, 246, 0.1)",
        border: "none",
        borderRadius: "4px",
        cursor: "pointer",
        padding: "4px",
        display: "flex",
        alignItems: "center",
        marginRight: "4px",
    },
    statsCard: {
        background: "var(--iccc-card-bg)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        padding: "12px 16px",
        marginTop: "auto",
        boxShadow: "var(--iccc-shadow)",
        display: "grid",
        gap: "6px",
    },
    statsRow: {
        display: "flex",
        justifyContent: "space-between",
        gap: "10px",
        fontSize: "11px",
        fontWeight: 600,
        color: "var(--iccc-text-muted)",
    },
    statsVal: {
        color: "var(--iccc-pill-active-bg)",
        fontWeight: 800,
        textAlign: "right",
        wordBreak: "break-word",
    },
};
