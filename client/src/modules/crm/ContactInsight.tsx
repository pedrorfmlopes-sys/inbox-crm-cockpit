import React from "react";
import * as Icons from "../../ui/icons";

interface ContactInsightProps {
    contact: {
        name: string;
        email: string;
        role?: string;
        company?: string;
        partnerLevel?: "Gold" | "Silver" | "Normal";
        salesVolume?: string;
        health?: "Great" | "Neutral" | "Risk";
    };
    onViewInOdoo?: () => void;
}

export const ContactInsight: React.FC<ContactInsightProps> = ({ contact, onViewInOdoo }) => {
    const getHealthColor = () => {
        switch (contact.health) {
            case "Great": return "#166534";
            case "Risk": return "#991b1b";
            default: return "#854d0e";
        }
    };

    const getHealthBg = () => {
        switch (contact.health) {
            case "Great": return "#dcfce7";
            case "Risk": return "#fee2e2";
            default: return "#fef9c3";
        }
    };

    return (
        <div style={S.container}>
            <div style={S.header}>
                <div style={S.avatar}>
                    {(contact.name?.[0] || contact.email?.[0] || "?").toUpperCase()}
                </div>
                <div style={S.info}>
                    <div style={S.name}>{contact.name || contact.email}</div>
                    <div style={S.roleCompany}>
                        {contact.role && <span>{contact.role}</span>}
                        {contact.role && contact.company && <span> @ </span>}
                        {contact.company && <span>{contact.company}</span>}
                    </div>
                </div>
                {contact.health && (
                    <div style={{ ...S.healthBadge, background: getHealthBg(), color: getHealthColor() }}>
                        {contact.health}
                    </div>
                )}
            </div>

            <div style={S.statsGrid}>
                <div style={S.statItem}>
                    <div style={S.statLabel}>Partner Level</div>
                    <div style={S.statValue}>
                        <Icons.Sparkles size={10} style={{ marginRight: '4px' }} />
                        {contact.partnerLevel || "Standard"}
                    </div>
                </div>
                <div style={S.statItem}>
                    <div style={S.statLabel}>Sales Volume</div>
                    <div style={S.statValue}>{contact.salesVolume || "€0"}</div>
                </div>
            </div>

            <div style={S.actions}>
                <button style={S.actionBtn} onClick={onViewInOdoo}>
                    <Icons.ExternalLink size={12} />
                    Odoo
                </button>
                <button style={S.actionBtn} onClick={async () => {
                    const { downloadVCard } = await import("./vCardService");
                    downloadVCard(contact as any);
                }}>
                    <Icons.Download size={12} />
                    vCard
                </button>
                <button style={S.actionBtn} onClick={async () => {
                    const { copyVisualCardToClipboard } = await import("./vCardService");
                    const ok = await copyVisualCardToClipboard(contact as any);
                    if (ok) alert("Cartão Visual copiado (Rich Text)!");
                }}>
                    <Icons.Clipboard size={12} title="Bypass IT attachment restriction" />
                    Copiar
                </button>
            </div>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    container: {
        padding: "10px",
        background: "var(--iccc-card-bg)",
        borderBottom: "1px solid var(--iccc-card-border)",
        position: "sticky",
        top: 0,
        zIndex: 10,
        boxShadow: "0 2px 8px rgba(0,0,0,0.05)",
    },
    header: {
        display: "flex",
        alignItems: "center",
        gap: "10px",
        marginBottom: "8px",
    },
    avatar: {
        width: "32px",
        height: "32px",
        borderRadius: "50%",
        background: "linear-gradient(135deg, #3b82f6 0%, #1d4ed8 100%)",
        color: "white",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        fontWeight: 700,
        fontSize: "14px",
    },
    info: {
        flex: 1,
        minWidth: 0,
    },
    name: {
        fontWeight: 700,
        fontSize: "13px",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
    },
    roleCompany: {
        fontSize: "11px",
        color: "var(--iccc-text-muted)",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
    },
    healthBadge: {
        fontSize: "9px",
        fontWeight: 800,
        padding: "2px 6px",
        borderRadius: "4px",
        textTransform: "uppercase",
    },
    statsGrid: {
        display: "grid",
        gridTemplateColumns: "1fr 1fr",
        gap: "8px",
        marginBottom: "8px",
    },
    statItem: {
        display: "flex",
        flexDirection: "column",
    },
    statLabel: {
        fontSize: "9px",
        color: "var(--iccc-text-muted)",
        textTransform: "uppercase",
        fontWeight: 700,
    },
    statValue: {
        fontSize: "11px",
        fontWeight: 600,
        display: "flex",
        alignItems: "center",
    },
    actions: {
        display: "flex",
        gap: "8px",
    },
    actionBtn: {
        flex: 1,
        background: "rgba(59, 130, 246, 0.05)",
        border: "1px solid rgba(59, 130, 246, 0.2)",
        borderRadius: "4px",
        padding: "4px 8px",
        fontSize: "11px",
        fontWeight: 600,
        color: "#2563eb",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "6px",
        cursor: "pointer",
    }
};
