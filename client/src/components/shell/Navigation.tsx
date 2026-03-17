import React from "react";
import { CockpitTab, useCockpit } from "@/components/shell/CockpitProvider";
import * as Icons from "../../ui/icons";

export const Navigation: React.FC = () => {
    const { tab, setTab, connectionStatus, granularStatusString, openSettingsSection } = useCockpit();

    const tabs: { id: CockpitTab; label: string; icon: React.ReactNode }[] = [
        { id: "ai", label: "AI", icon: <Icons.Sparkles size={16} /> },
        { id: "crm", label: "CRM", icon: <Icons.Database size={16} /> },
        { id: "crm2", label: "CRM 2", icon: <Icons.Database size={16} /> },
        { id: "related", label: "Contexto", icon: <Icons.Link size={16} /> },
        { id: "groups", label: "Grupos", icon: <Icons.Clipboard size={16} /> },
        { id: "files", label: "Files", icon: <Icons.Files size={16} /> },
        { id: "settings", label: "Settings", icon: <Icons.Settings size={16} /> },
    ];

    return (
        <div style={S.navWrapper}>
            {tabs.map((t) => {
                const isSettings = t.id === "settings";
                const statusColor = connectionStatus === "success" ? "#36b37e" : "#ff5630";
                return (
                    <div key={t.id} style={S.tabSlot}>
                        <button
                            style={tab === t.id ? S.tabActive : S.tab}
                            onClick={() => setTab(t.id)}
                        >
                            <span style={S.icon}>{t.icon}</span>
                            <span style={S.label}>{t.label}</span>
                        </button>

                        {isSettings && connectionStatus !== "none" && (
                            <button
                                type="button"
                                title={`Abrir Ligações. ${granularStatusString}`}
                                aria-label="Abrir ligações"
                                style={{
                                    ...S.settingsStatusDotBtn,
                                    background: statusColor,
                                }}
                                onClick={(event) => {
                                    event.stopPropagation();
                                    openSettingsSection("conns");
                                }}
                            />
                        )}
                    </div>
                );
            })}
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
    tabSlot: {
        flex: 1,
        position: "relative",
        display: "flex",
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
    settingsStatusDotBtn: {
        position: "absolute",
        top: "8px",
        right: "14px",
        width: "11px",
        height: "11px",
        borderRadius: "50%",
        border: "2px solid var(--iccc-card-bg)",
        boxShadow: "0 2px 6px rgba(0,0,0,0.18)",
        cursor: "pointer",
        padding: 0,
    },
};
