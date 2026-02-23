import React, { useState } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";

export const GlobalHeader: React.FC = () => {
    const cockpit = useCockpit();
    if (!cockpit) return null;
    const { ctx, logout } = cockpit;
    const [isExpanded, setIsExpanded] = useState(false);

    return (
        <div style={S.header}>
            <div style={S.topRow}>
                <div style={S.subjectBlock}>
                    <div style={S.label}>Assunto</div>
                    <div style={S.subject} title={ctx.subject || ""}>
                        {ctx.subject || "Sem assunto"}
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
