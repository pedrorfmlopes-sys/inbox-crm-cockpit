import React, { useState } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";

export const GlobalHeader: React.FC = () => {
    const [isExpanded, setIsExpanded] = useState(false);
    const { ctx, links, logout } = useCockpit();
    const hasActiveLinks = links.length > 0;

    return (
        <div style={S.header}>
            <div style={S.topRow}>
                <div style={S.subjectBlock}>
                    <div style={S.label}>Assunto</div>
                    <div style={S.subjectRow}>
                        <div style={S.subject} title={ctx.subject || ""}>
                            {ctx.subject || "Sem assunto"}
                        </div>
                        {hasActiveLinks ? (
                            <div style={S.linkedBadge} title="Este email tem ligacao ativa ao Odoo">
                                Odoo ligado · {links.length}
                            </div>
                        ) : null}
                    </div>
                </div>
                <div style={{ display: "flex", gap: "8px" }}>
                    <button style={S.expandBtn} onClick={logout}>
                        Sair
                    </button>
                    <button style={S.expandBtn} onClick={() => setIsExpanded(!isExpanded)}>
                        {isExpanded ? "▲" : "▼"}
                    </button>
                </div>
            </div>

            {isExpanded && (
                <div style={S.details}>
                    {hasActiveLinks ? (
                        <div style={S.linkedCard}>
                            <div style={S.linkedCardTitle}>Ligacao ativa ao Odoo</div>
                            <div style={S.linkedCardText}>
                                Este email ja esta ligado a {links.length} registo{links.length > 1 ? "s" : ""}.
                            </div>
                        </div>
                    ) : null}
                    <div style={S.row}>
                        <div style={S.label}>De</div>
                        <div style={S.value}>{ctx.fromName ? `${ctx.fromName} <${ctx.fromEmail}>` : ctx.fromEmail || "—"}</div>
                    </div>
                    <div style={S.row}>
                        <div style={S.label}>Thread</div>
                        <div style={S.monoValue}>{ctx.conversationId || "—"}</div>
                    </div>
                </div>
            )}
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    header: {
        margin: "8px 12px 12px 12px",
        padding: "10px 14px",
        background: "var(--iccc-card-bg)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "var(--iccc-radius-card)",
        boxShadow: "var(--iccc-shadow)",
        backdropFilter: "var(--iccc-glass-blur)",
        WebkitBackdropFilter: "var(--iccc-glass-blur)",
    },
    topRow: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "flex-start",
        gap: "12px",
    },
    subjectBlock: {
        minWidth: 0,
        flex: 1,
    },
    subjectRow: {
        display: "flex",
        alignItems: "center",
        gap: "8px",
        minWidth: 0,
        flexWrap: "wrap",
    },
    label: {
        fontSize: "9px",
        fontWeight: 600,
        textTransform: "uppercase",
        letterSpacing: "0.05em",
        color: "var(--iccc-text-muted)",
        marginBottom: "4px",
    },
    subject: {
        fontSize: "13px",
        fontWeight: 700,
        color: "var(--iccc-text)",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
        minWidth: 0,
        flex: 1,
    },
    linkedBadge: {
        fontSize: "10px",
        fontWeight: 700,
        color: "#006644",
        background: "#E3FCEF",
        border: "1px solid #ABF5D1",
        borderRadius: "999px",
        padding: "3px 8px",
        whiteSpace: "nowrap",
    },
    expandBtn: {
        padding: "4px 8px",
        fontSize: "10px",
        fontWeight: 600,
        background: "rgba(0,0,0,0.05)",
        border: "none",
        borderRadius: "6px",
        cursor: "pointer",
        color: "var(--iccc-text)",
    },
    details: {
        marginTop: "12px",
        paddingTop: "12px",
        borderTop: "1px solid rgba(0,0,0,0.05)",
        display: "flex",
        flexDirection: "column",
        gap: "8px",
    },
    linkedCard: {
        border: "1px solid #ABF5D1",
        background: "#E3FCEF",
        borderRadius: "8px",
        padding: "8px 10px",
    },
    linkedCardTitle: {
        fontSize: "11px",
        fontWeight: 700,
        color: "#006644",
        marginBottom: "2px",
    },
    linkedCardText: {
        fontSize: "11px",
        color: "#006644",
    },
    row: {
        display: "flex",
        flexDirection: "column",
    },
    value: {
        fontSize: "11px",
        fontWeight: 500,
        color: "var(--iccc-text)",
    },
    monoValue: {
        fontSize: "10px",
        fontFamily: "ui-monospace, SFMono-Regular, Menlo, Monaco, Consolas, monospace",
        color: "var(--iccc-text-muted)",
        wordBreak: "break-all",
    },
};
