import React from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { openCockpitDialog } from "../../office";
import * as Icons from "../../ui/icons";

export const CrmCockpit: React.FC = () => {
    const { ctx, meta, links, msg, refreshLinks, setMsg } = useCockpit();

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
                ...(extra || {}),
            });
            await refreshLinks();
        } catch (e: any) {
            setMsg(e?.message ?? String(e));
        }
    }

    return (
        <div style={S.container}>
            <div style={S.actionRow}>
                <button style={S.primaryBtn} onClick={() => openDialog("new")}>
                    <Icons.Plus size={16} />
                    Criar Item
                </button>
                <button style={S.secondaryBtn} onClick={() => openDialog("add")}>
                    <Icons.Link size={16} />
                    Ligar Existente
                </button>
            </div>

            {msg && <div style={S.alert}>{msg}</div>}

            <div style={S.section}>
                <div style={S.sectionHeader}>
                    <h3 style={S.sectionTitle}>Ligados a esta conversa</h3>
                    <button style={S.refreshBtn} onClick={() => refreshLinks()}>
                        <Icons.RefreshCw size={14} />
                    </button>
                </div>

                {!links.length ? (
                    <div style={S.emptyState}>
                        <div style={S.emptyIcon}>
                            <Icons.Files size={32} />
                        </div>
                        <p>{!meta ? "Odoo não configurado no servidor." : "Nenhum registo do Odoo associado."}</p>
                        {!meta && <p style={{ fontSize: '11px', marginTop: '8px', opacity: 0.7 }}>Define as tuas credenciais no ficheiro .env para ligar emails ao CRM.</p>}
                    </div>
                ) : (
                    <div style={S.cardList}>
                        {links.map((link) => (
                            <JiraCard
                                key={`${link.model}-${link.recordId}`}
                                link={link}
                                meta={meta}
                                onEdit={() => openDialog("edit", { model: link.model, recordId: String(link.recordId) })}
                            />
                        ))}
                    </div>
                )}
            </div>

            {meta && (
                <div style={S.footer}>
                    Conectado a: <strong>{meta.baseUrl}</strong> ({meta.db})
                </div>
            )}
        </div>
    );
};

const JiraCard: React.FC<{ link: any; meta: any; onEdit: () => void }> = ({ link, meta, onEdit }) => {
    const modelPrefix = link.model.split(".")[0] || link.model;

    // Odoo URL logic
    const base = meta?.baseUrl || meta?.webBaseUrl || meta?.url || "";
    const url = base
        ? `${String(base).replace(/\/+$/, "")}/web#id=${link.recordId}&model=${encodeURIComponent(link.model)}&view_type=form`
        : "";

    const getStatusStyle = (model: string) => {
        if (model.includes("project")) return { bg: "#dbeafe", color: "#1e40af", label: "Em curso" };
        if (model.includes("task")) return { bg: "#fef9c3", color: "#854d0e", label: "Pendente" };
        return { bg: "#dcfce7", color: "#166534", label: "Ativo" };
    };

    const status = getStatusStyle(link.model);

    const copyToClipboard = () => {
        if (url) {
            navigator.clipboard.writeText(url);
            alert("Link copiado para a área de transferência!");
        }
    };

    return (
        <div style={S.card}>
            <div style={S.cardHeader}>
                <span style={S.modelTag}>{modelPrefix.toUpperCase()}</span>
                <span style={S.recordId}>#{link.recordId}</span>
                <div style={{ flex: 1 }} />
                <span style={{ ...S.statusPill, background: status.bg, color: status.color }}>
                    {status.label}
                </span>
            </div>

            <div style={S.cardTitle}>{link.recordName || link.title || "Sem título"}</div>

            <div style={S.cardFooter}>
                <div style={S.cardActions}>
                    <button style={S.cardActionBtn} onClick={onEdit}>
                        <Icons.Edit size={12} style={{ marginRight: "4px" }} />
                        Editar
                    </button>
                    <button style={S.cardActionBtn} onClick={copyToClipboard}>
                        <Icons.Clipboard size={12} style={{ marginRight: "4px" }} />
                        Copiar Link
                    </button>
                    {url && (
                        <a href={url} target="_blank" rel="noreferrer" style={S.cardActionBtn}>
                            <Icons.ExternalLink size={12} style={{ marginRight: "4px" }} />
                            Ver no Odoo
                        </a>
                    )}
                </div>
            </div>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    container: {
        display: "flex",
        flexDirection: "column",
        gap: "20px",
        paddingTop: "4px",
    },
    actionRow: {
        display: "flex",
        gap: "10px",
    },
    primaryBtn: {
        flex: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "6px",
        padding: "12px",
        background: "var(--iccc-btn-bg)",
        color: "var(--iccc-btn-text)",
        border: "none",
        borderRadius: "var(--iccc-radius-btn)",
        fontWeight: 700,
        fontSize: "13px",
        cursor: "pointer",
        boxShadow: "0 4px 12px rgba(37, 99, 235, 0.2)",
    },
    secondaryBtn: {
        flex: 1,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "6px",
        padding: "12px",
        background: "var(--iccc-btn2-bg)",
        color: "var(--iccc-btn2-text)",
        border: "1px solid var(--iccc-btn2-border)",
        borderRadius: "var(--iccc-radius-btn)",
        fontWeight: 700,
        fontSize: "13px",
        cursor: "pointer",
    },
    btnIcon: {
        fontSize: "16px",
    },
    alert: {
        padding: "12px",
        background: "#fee2e2",
        color: "#991b1b",
        borderRadius: "12px",
        fontSize: "12px",
        fontWeight: 500,
    },
    section: {
        display: "flex",
        flexDirection: "column",
        gap: "12px",
    },
    sectionHeader: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
    },
    sectionTitle: {
        fontSize: "12px",
        fontWeight: 700,
        textTransform: "uppercase",
        letterSpacing: "0.05em",
        color: "var(--iccc-text-muted)",
        margin: 0,
    },
    refreshBtn: {
        background: "none",
        border: "none",
        color: "var(--iccc-text-muted)",
        fontSize: "16px",
        cursor: "pointer",
        padding: "4px",
    },
    emptyState: {
        padding: "32px 16px",
        textAlign: "center",
        background: "var(--iccc-card-bg)",
        border: "1px dashed var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        color: "var(--iccc-text-muted)",
    },
    emptyIcon: {
        fontSize: "24px",
        marginBottom: "8px",
        opacity: 0.5,
    },
    cardList: {
        display: "flex",
        flexDirection: "column",
        gap: "10px",
    },
    card: {
        background: "var(--iccc-card-bg)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        padding: "14px",
        boxShadow: "var(--iccc-shadow)",
        backdropFilter: "var(--iccc-glass-blur)",
        WebkitBackdropFilter: "var(--iccc-glass-blur)",
        transition: "transform 0.2s ease",
    },
    cardHeader: {
        display: "flex",
        alignItems: "center",
        gap: "8px",
        marginBottom: "8px",
    },
    modelTag: {
        fontSize: "9px",
        fontWeight: 800,
        padding: "2px 6px",
        background: "rgba(37, 99, 235, 0.1)",
        color: "#2563eb",
        borderRadius: "4px",
    },
    recordId: {
        fontSize: "10px",
        fontWeight: 600,
        color: "var(--iccc-text-muted)",
    },
    statusPill: {
        fontSize: "9px",
        fontWeight: 700,
        padding: "2px 8px",
        background: "#dcfce7",
        color: "#166534",
        borderRadius: "999px",
    },
    cardTitle: {
        fontSize: "14px",
        fontWeight: 600,
        color: "var(--iccc-text)",
        marginBottom: "12px",
        lineHeight: "1.4",
    },
    cardFooter: {
        display: "flex",
        justifyContent: "flex-end",
        borderTop: "1px solid rgba(0,0,0,0.03)",
        paddingTop: "10px",
    },
    cardActions: {
        display: "flex",
        gap: "12px",
    },
    cardActionBtn: {
        background: "none",
        border: "none",
        color: "#2563eb",
        fontSize: "11px",
        fontWeight: 700,
        cursor: "pointer",
        padding: 0,
        textDecoration: "none",
    },
    footer: {
        marginTop: "auto",
        padding: "12px",
        fontSize: "10px",
        textAlign: "center",
        color: "var(--iccc-text-muted)",
        borderTop: "1px solid var(--iccc-card-border)",
    },
};
