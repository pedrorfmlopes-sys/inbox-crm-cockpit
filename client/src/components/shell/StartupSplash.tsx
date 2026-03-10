import React from "react";
import type { StartupCheck } from "./CockpitProvider";
import * as Icons from "@/ui/icons";

export function StartupSplash({ checks }: { checks: StartupCheck[] }): JSX.Element {
    const completed = checks.filter((check) => check.status === "success").length;
    const attention = checks.filter((check) => check.status === "warning" || check.status === "error").length;

    return (
        <div style={S.root}>
            <div style={S.backgroundGlow} />
            <div style={S.content}>
                <div style={S.hero}>
                    <div style={S.brandLockup}>
                        <div style={S.logoShell}>
                            <div style={S.logoHalo} />
                            <Icons.Lock size={28} color="#F8FAFC" />
                            <div style={S.sparkle}>
                                <Icons.Sparkles size={14} color="#93C5FD" />
                            </div>
                        </div>
                        <div style={S.brandText}>
                            <div style={S.brandEyebrow}>Inbox CRM Cockpit</div>
                            <div style={S.brandTitle}>A preparar o teu cockpit</div>
                            <div style={S.brandSubtitle}>
                                Verificamos sessão, email atual, ligações e serviços antes de abrir a aplicação.
                            </div>
                        </div>
                    </div>

                    <div style={S.heroStats}>
                        <div style={S.statPill}>
                            <Icons.Check size={14} color="#0F766E" />
                            <span>{completed}/{checks.length} prontos</span>
                        </div>
                        <div style={S.statPill}>
                            <Icons.Database size={14} color="#1D4ED8" />
                            <span>{attention ? `${attention} alerta(s)` : "Tudo dentro do esperado"}</span>
                        </div>
                    </div>
                </div>

                <div style={S.panel}>
                    {checks.map((check) => {
                        const palette = getPalette(check.status);
                        return (
                            <div key={check.id} style={{ ...S.checkRow, borderColor: palette.border, background: palette.background }}>
                                <div style={{ ...S.checkIcon, background: palette.badge }}>
                                    {renderStatusIcon(check.status, palette.icon)}
                                </div>
                                <div style={S.checkBody}>
                                    <div style={S.checkLabel}>{check.label}</div>
                                    <div style={{ ...S.checkDetail, color: palette.text }}>{check.detail}</div>
                                </div>
                            </div>
                        );
                    })}
                </div>
            </div>

            <style>{`
                @keyframes iccc-startup-pulse {
                    0%, 100% { transform: scale(1); opacity: .65; }
                    50% { transform: scale(1.08); opacity: 1; }
                }
                @keyframes iccc-startup-float {
                    0%, 100% { transform: translateY(0px); }
                    50% { transform: translateY(-6px); }
                }
                @keyframes iccc-startup-spin {
                    to { transform: rotate(360deg); }
                }
            `}</style>
        </div>
    );
}

function renderStatusIcon(status: StartupCheck["status"], color: string): JSX.Element {
    if (status === "success") return <Icons.Check size={14} color={color} />;
    if (status === "warning" || status === "error") return <Icons.AlertTriangle size={14} color={color} />;
    if (status === "running") return <Icons.RefreshCw size={14} color={color} style={{ animation: "iccc-startup-spin 1s linear infinite" }} />;
    return <Icons.Clock size={14} color={color} />;
}

function getPalette(status: StartupCheck["status"]) {
    if (status === "success") {
        return {
            background: "rgba(236, 253, 245, 0.85)",
            border: "rgba(16, 185, 129, 0.28)",
            badge: "rgba(16, 185, 129, 0.12)",
            icon: "#047857",
            text: "#065F46",
        };
    }
    if (status === "warning") {
        return {
            background: "rgba(255, 247, 237, 0.92)",
            border: "rgba(249, 115, 22, 0.25)",
            badge: "rgba(249, 115, 22, 0.12)",
            icon: "#C2410C",
            text: "#9A3412",
        };
    }
    if (status === "error") {
        return {
            background: "rgba(254, 242, 242, 0.92)",
            border: "rgba(239, 68, 68, 0.24)",
            badge: "rgba(239, 68, 68, 0.12)",
            icon: "#B91C1C",
            text: "#991B1B",
        };
    }
    if (status === "running") {
        return {
            background: "rgba(239, 246, 255, 0.92)",
            border: "rgba(59, 130, 246, 0.22)",
            badge: "rgba(59, 130, 246, 0.12)",
            icon: "#2563EB",
            text: "#1D4ED8",
        };
    }
    return {
        background: "rgba(248, 250, 252, 0.92)",
        border: "rgba(148, 163, 184, 0.2)",
        badge: "rgba(148, 163, 184, 0.12)",
        icon: "#475569",
        text: "#64748B",
    };
}

const S: Record<string, React.CSSProperties> = {
    root: {
        position: "relative",
        minHeight: "100%",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        overflow: "hidden",
        padding: "24px 16px",
        background: "linear-gradient(180deg, #E8EEF7 0%, #F7FAFC 48%, #E6EEF5 100%)",
    },
    backgroundGlow: {
        position: "absolute",
        inset: "auto auto 6% -18%",
        width: 260,
        height: 260,
        borderRadius: "50%",
        background: "radial-gradient(circle, rgba(59,130,246,.18) 0%, rgba(59,130,246,0) 72%)",
        animation: "iccc-startup-pulse 4s ease-in-out infinite",
        pointerEvents: "none",
    },
    content: {
        position: "relative",
        zIndex: 1,
        display: "grid",
        gap: 18,
        width: "100%",
        maxWidth: 420,
    },
    hero: {
        display: "grid",
        gap: 14,
    },
    brandLockup: {
        display: "grid",
        gap: 14,
        gridTemplateColumns: "minmax(0, 84px) minmax(0, 1fr)",
        alignItems: "center",
    },
    logoShell: {
        position: "relative",
        width: 84,
        height: 84,
        borderRadius: 28,
        background: "linear-gradient(145deg, #0F172A 0%, #1D4ED8 55%, #38BDF8 100%)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        boxShadow: "0 26px 46px rgba(15, 23, 42, 0.18)",
        animation: "iccc-startup-float 4.6s ease-in-out infinite",
    },
    logoHalo: {
        position: "absolute",
        inset: -10,
        borderRadius: 36,
        border: "1px solid rgba(59, 130, 246, 0.18)",
        background: "radial-gradient(circle at 30% 30%, rgba(255,255,255,.22), rgba(255,255,255,0))",
    },
    sparkle: {
        position: "absolute",
        right: 10,
        top: 10,
        width: 24,
        height: 24,
        borderRadius: 999,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        background: "rgba(15, 23, 42, 0.45)",
        backdropFilter: "blur(10px)",
    },
    brandText: {
        display: "grid",
        gap: 4,
        minWidth: 0,
    },
    brandEyebrow: {
        fontSize: 11,
        fontWeight: 800,
        textTransform: "uppercase",
        letterSpacing: 1.3,
        color: "#2563EB",
    },
    brandTitle: {
        fontSize: 24,
        fontWeight: 800,
        lineHeight: 1.05,
        color: "#0F172A",
    },
    brandSubtitle: {
        fontSize: 13,
        lineHeight: 1.5,
        color: "#475569",
    },
    heroStats: {
        display: "flex",
        gap: 8,
        flexWrap: "wrap",
    },
    statPill: {
        display: "inline-flex",
        alignItems: "center",
        gap: 6,
        padding: "8px 10px",
        borderRadius: 999,
        background: "rgba(255,255,255,.72)",
        border: "1px solid rgba(148,163,184,.24)",
        color: "#0F172A",
        fontSize: 12,
        fontWeight: 700,
        backdropFilter: "blur(12px)",
    },
    panel: {
        display: "grid",
        gap: 10,
        padding: 14,
        borderRadius: 20,
        background: "rgba(255,255,255,.72)",
        border: "1px solid rgba(148, 163, 184, 0.16)",
        boxShadow: "0 16px 28px rgba(15, 23, 42, 0.08)",
        backdropFilter: "blur(18px)",
    },
    checkRow: {
        display: "grid",
        gridTemplateColumns: "36px minmax(0, 1fr)",
        alignItems: "center",
        gap: 12,
        borderRadius: 14,
        border: "1px solid transparent",
        padding: "10px 12px",
        minWidth: 0,
    },
    checkIcon: {
        width: 36,
        height: 36,
        borderRadius: 12,
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        flexShrink: 0,
    },
    checkBody: {
        display: "grid",
        gap: 2,
        minWidth: 0,
    },
    checkLabel: {
        fontSize: 12,
        fontWeight: 800,
        color: "#0F172A",
    },
    checkDetail: {
        fontSize: 12,
        lineHeight: 1.45,
        wordBreak: "break-word",
    },
};
