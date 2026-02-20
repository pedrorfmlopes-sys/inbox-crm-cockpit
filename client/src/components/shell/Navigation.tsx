import React from "react";
import { CockpitTab, useCockpit } from "@/components/shell/CockpitProvider";
import * as Icons from "../../ui/icons";

export const Navigation: React.FC = () => {
    const { tab, setTab } = useCockpit();

    const tabs: { id: CockpitTab; label: string; icon: React.ReactNode }[] = [
        { id: "ai", label: "AI", icon: <Icons.Sparkles size={16} /> },
        { id: "crm", label: "CRM", icon: <Icons.Database size={16} /> },
        { id: "files", label: "Files", icon: <Icons.Files size={16} /> },
        { id: "settings", label: "Settings", icon: <Icons.Settings size={16} /> },
    ];

    return (
        <div style={S.navWrapper}>
            <div style={S.nav}>
                {tabs.map((t) => (
                    <button
                        key={t.id}
                        style={tab === t.id ? S.tabActive : S.tab}
                        onClick={() => setTab(t.id)}
                    >
                        <span style={S.icon}>{t.icon}</span>
                        <span style={S.label}>{t.label}</span>
                    </button>
                ))}
            </div>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    navWrapper: {
        position: "fixed",
        bottom: "16px",
        left: "0",
        right: "0",
        display: "flex",
        justifyContent: "center",
        padding: "0 16px",
        zIndex: 1000,
        pointerEvents: "none",
    },
    nav: {
        display: "flex",
        background: "var(--iccc-bottom-bg)",
        backdropFilter: "blur(8px)",
        WebkitBackdropFilter: "blur(8px)",
        border: "1px solid var(--iccc-bottom-border)",
        borderRadius: "var(--iccc-bottom-radius)",
        padding: "4px",
        gap: "2px",
        boxShadow: "var(--iccc-bottom-shadow)",
        pointerEvents: "auto",
    },
    tab: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: "4px 8px",
        minWidth: "56px",
        borderRadius: "10px",
        border: "none",
        background: "transparent",
        cursor: "pointer",
        transition: "all 0.2s ease",
        color: "var(--iccc-text-muted)",
    },
    tabActive: {
        display: "flex",
        flexDirection: "column",
        alignItems: "center",
        justifyContent: "center",
        padding: "4px 8px",
        minWidth: "56px",
        borderRadius: "10px",
        border: "none",
        background: "var(--iccc-pill-active-bg)",
        cursor: "pointer",
        transition: "all 0.2s ease",
        color: "var(--iccc-pill-active-text)",
    },
    icon: {
        fontSize: "18px",
        marginBottom: "2px",
    },
    label: {
        fontSize: "9px",
        fontWeight: 600,
        textTransform: "uppercase",
        letterSpacing: "0.01em",
    },
};
