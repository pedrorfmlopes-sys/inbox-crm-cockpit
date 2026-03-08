import React, { useState, useRef } from "react";
import * as Icons from "../../ui/icons";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { getAttachments } from "@/office"; // We'll assume this is exported now
import { PanelState, type PanelStateTone } from "../../ui/PanelState";

export const FileCockpit: React.FC = () => {
    const { files, addFile, removeFile, setMsg, setTab, setAiState } = useCockpit();
    const [isDragging, setIsDragging] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [status, setStatus] = useState<{ tone: PanelStateTone; title: string; description?: string } | null>(null);
    const hiddenFileInput = useRef<HTMLInputElement>(null);

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
                const base64 = (reader.result as string).split(",")[1]; // Remove data:application/pdf;base64, prefix
                addFile({
                    name: file.name,
                    type: file.type,
                    content: base64
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
                description: `${added} ficheiro(s) prontos para análise pela IA.`,
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
        setAiState({ prompt: `ANÁLISE DETALHADA DO ANEXO "${fileName}": Extrai datas, valores, nº faturas e descreve o conteúdo técnico.` });
        setTab("ai");
    };

    const handleImportAttachments = async () => {
        try {
            setIsImporting(true);
            setStatus({
                tone: "loading",
                title: "A importar anexos",
                description: "Estamos a ler os anexos disponíveis no email atual.",
            });
            const atts = await getAttachments();
            if (atts.length === 0) {
                setMsg("Nenhum anexo encontrado neste email.");
                setStatus({
                    tone: "empty",
                    title: "Sem anexos no email",
                    description: "Podes importar ficheiros do computador ou selecionar outro email com anexos.",
                });
            } else {
                let counts = 0;
                // We need to fetch content. The getAttachments in office.ts assumes it fetches content.
                // But wait, the standard getAttachments usually returns metadata. 
                // We need to ensure getAttachments fetches the base64 content.
                // My previous implementation of getAttachments DOES fetch content.

                // However, getAttachments() implementation in office.ts I just reviewed seems to actually 
                // iterate and call getAttachmentContentAsync, so it should return content.
                // Let's verify the implementation I just wrote... Yes, it returns { content }.

                // But wait, getAttachments implementation I wrote logic:
                // It iterates attachments and calls getAttachmentContentAsync.
                // So the result is full content.

                for (const att of atts) {
                    addFile({
                        name: att.name,
                        type: att.contentType,
                        content: att.content // it is already base64 from getAttachmentContentAsync usually (if format is base64)
                    });
                    counts++;
                }
                setMsg(`${counts} anexos importados com sucesso!`);
                setStatus({
                    tone: "success",
                    title: "Anexos importados",
                    description: `${counts} anexo(s) disponíveis para leitura e resumo.`,
                });
            }
        } catch (e: any) {
            console.error("Import error", e);
            setMsg("Erro ao importar anexos: " + e.message);
            setStatus({
                tone: "error",
                title: "Falha ao importar anexos",
                description: e?.message || "Não foi possível ler os anexos deste email.",
            });
        } finally {
            setIsImporting(false);
        }
    };

    return (
        <div style={S.container}>
            <div style={S.header}>
                <h3 style={S.title}>Documentos & Anexos</h3>
                <p style={S.subtitle}>Ficheiros carregados serão lidos pela AI</p>
            </div>

            {status && (
                <PanelState
                    tone={status.tone}
                    title={status.title}
                    description={status.description}
                    compact
                />
            )}

            {/* ACTION BUTTONS */}
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

            {/* DROP ZONE */}
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

            {/* FILE LIST */}
            <div style={S.fileList}>
                {files.length === 0 && (
                    <PanelState
                        tone="empty"
                        title="Nenhum ficheiro carregado"
                        description="Importa anexos do email ou adiciona ficheiros do computador para os analisar."
                    />
                )}

                {files.map((f, idx) => (
                    <div key={idx} style={S.fileItem}>
                        <div style={S.fileIcon}>
                            {f.type.includes("pdf") ? <Icons.Receipt size={16} color="#ef4444" /> : <Icons.Files size={16} color="#3b82f6" />}
                        </div>
                        <div style={S.fileInfo}>
                            <div style={S.fileName} title={f.name}>{f.name}</div>
                            <div style={S.fileMeta}>{f.type || "unknown"}</div>
                        </div>
                        <button
                            style={S.actionIconBtn}
                            onClick={() => triggerFileSummary(f.name)}
                            title="Resumir este ficheiro na AI"
                        >
                            <Icons.Sparkles size={14} color="var(--iccc-pill-active-bg)" />
                        </button>
                        <button style={S.deleteBtn} onClick={() => removeFile(f.name)} title="Remover">
                            ✕
                        </button>
                    </div>
                ))}
            </div>

            <div style={S.statsCard}>
                <div style={S.statsRow}>
                    <span>Ficheiros em Memória</span>
                    <span style={S.statsVal}>{files.length}</span>
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
    dropZone: {
        border: "2px dashed var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        padding: "24px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        transition: "all 0.2s ease",
        cursor: "default", // Drag area
        minHeight: "100px",
    },
    fileList: {
        display: "flex",
        flexDirection: "column",
        gap: "8px",
        maxHeight: "200px",
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
    },
    statsRow: {
        display: "flex",
        justifyContent: "space-between",
        fontSize: "11px",
        fontWeight: 600,
        color: "var(--iccc-text-muted)",
    },
    statsVal: {
        color: "var(--iccc-pill-active-bg)",
        fontWeight: 800,
    },
};
