import React from "react";
import { CockpitTab, useCockpit } from "@/components/shell/CockpitProvider";
import * as Icons from "../../ui/icons";

export const Navigation: React.FC = () => {
    const { tab, setTab, connectionStatus, granularStatusString } = useCockpit();

    const tabs: { id: CockpitTab; label: string; icon: React.ReactNode }[] = [
        { id: "ai", label: "AI", icon: <Icons.Sparkles size={16} /> },
        { id: "crm", label: "CRM", icon: <Icons.Database size={16} /> },
        { id: "related", label: "Contexto", icon: <Icons.Link size={16} /> },
        { id: "files", label: "Files", icon: <Icons.Files size={16} /> },
        { id: "settings", label: "Settings", icon: <Icons.Settings size={16} /> },
    ];

    return (
        <div style={S.navWrapper}>
            {tabs.map((t) => (
                <button
                    key={t.id}
                    style={tab === t.id ? S.tabActive : S.tab}
                    onClick={() => setTab(t.id)}
                >
                    <span style={S.icon}>
                        {t.icon}
                        {t.id === "crm" && connectionStatus !== "none" && (
                            <div
                                title={granularStatusString}
                                style={{
                                    ...S.statusDot,
                                    background: connectionStatus === "success" ? "#36b37e" : "#ff5630"
                                }}
                            />
                        )}
                    </span>
                    <span style={S.label}>{t.label}</span>
                </button>
            ))}
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    navWrapper: {
        position: "absolute",
        bottom: "0",
        left: "0",
        right: "0",
        height: "60px",
        display: "flex",
        background: "var(--iccc-card-bg)",
        borderTop: "1px solid var(--iccc-card-border)",
        boxShadow: "0 -2px 10px rgba(0,0,0,0.05)",
        zIndex: 1000,
    },
    tab: {
        flex: 1,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: "6px 0",
        border: "none",
        borderTop: "2px solid transparent",
        background: "transparent",
        cursor: "pointer",
        transition: "all 0.2s ease",
        color: "var(--iccc-text-muted)",
    },
    tabActive: {
        flex: 1,
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: "6px 0",
        border: "none",
        borderTop: "2px solid var(--iccc-pill-active-bg)",
        background: "rgba(59, 130, 246, 0.05)",
        cursor: "pointer",
        transition: "all 0.2s ease",
        color: "var(--iccc-pill-active-bg)",
    },
    icon: {
        fontSize: "18px",
        marginBottom: "2px",
        position: "relative",
        display: "inline-flex",
    },
    label: {
        fontSize: "9px",
        fontWeight: 600,
        textTransform: "uppercase",
        letterSpacing: "0.01em",
    },
    statusDot: {
        position: "absolute",
        top: "-2px",
        right: "-4px",
        width: "6px",
        height: "6px",
        borderRadius: "50%",
        border: "1px solid white",
    },
};
