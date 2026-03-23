import React from "react";
import { CockpitProvider, useCockpit } from "./CockpitProvider";
import { Navigation } from "./Navigation";
import { GlobalHeader } from "./GlobalHeader";
import { StartupSplash } from "./StartupSplash";
import { PanelState } from "@/ui/PanelState";

// Modules
import { AiCockpit } from "../../modules/ai/AiCockpit";
import { CrmCockpit } from "../../modules/crm/CrmCockpit";
import { CrmCockpit2 } from "../../modules/crm/CrmCockpit2";
import { RelatedCockpit } from "../../modules/crm/RelatedCockpit";
import { GroupManagerCockpit } from "../../modules/crm/GroupManagerCockpit";
import { FileCockpit } from "../../modules/files/FileCockpit";
import { SettingsPanel } from "../../ui/SettingsPanel";

import { LoginCockpit } from "../../modules/auth/LoginCockpit";

function ShellContent() {
    const cockpit = useCockpit();
    if (!cockpit) return null;
    const { tab, isAuthenticated, isLoading, startupChecks, startupNotice, dismissStartupNotice, settings, openSettingsSection } = cockpit as any;
    const hasOdooSetup = Boolean(
        String(settings?.odooUrl || "").trim()
        || String(settings?.odooDb || "").trim()
        || String(settings?.odooLogin || "").trim()
        || String(settings?.odooPassword || "").trim()
        || String(settings?.odooSessionToken || "").trim()
    );
    const tabRequiresOdoo = tab === "crm" || tab === "crm2" || tab === "related";

    if (isLoading) {
        return <StartupSplash checks={startupChecks} />;
    }

    return (
        <div style={{
            display: "flex",
            flexDirection: "column",
            height: "100%",
            background: "var(--iccc-bg)",
            color: "var(--iccc-text)",
            fontFamily: "var(--iccc-font)",
            overflow: "hidden",
            position: "relative"
        }}>
            <GlobalHeader />

            <main className="flex-1 relative" style={{
                padding: "10px",
                overflowY: "scroll",
                flex: "1 1 0%",
                minHeight: 0
            }}>
                {startupNotice ? <StartupNoticeBanner notice={startupNotice} onDismiss={dismissStartupNotice} /> : null}
                {tabRequiresOdoo && !isAuthenticated ? (
                    <div style={S.gateRoot}>
                        {hasOdooSetup ? (
                            <LoginCockpit />
                        ) : (
                            <div style={S.odooGateCard}>
                                <PanelState
                                    compact
                                    tone="info"
                                    title="Integracao Odoo nao configurada"
                                    description="Esta area depende do Odoo. Podes continuar a usar Grupos, Files, AI e Settings sem Odoo."
                                />
                                <button type="button" style={S.odooGateBtn} onClick={() => openSettingsSection("conns")}>
                                    Abrir Ligacoes
                                </button>
                            </div>
                        )}
                    </div>
                ) : (
                    <>
                        {tab === "ai" && <AiCockpit />}
                        {tab === "crm" && <CrmCockpit />}
                        {tab === "crm2" && <CrmCockpit2 />}
                        {tab === "related" && <RelatedCockpit />}
                        {tab === "groups" && <GroupManagerCockpit />}
                        {tab === "files" && <FileCockpit />}
                        {tab === "settings" && <SettingsPanel />}
                    </>
                )}
            </main>

            {/* Ghost Spacer: Faz com que o 'main' pare exatamente no topo da barra, 
                ativando o scroll lateral sem tapar o chat. */}
            <div style={{ height: "60px", flexShrink: 0 }} />

            <Navigation />
        </div>
    );
}

export function CockpitShell() {
    return (
        <CockpitProvider>
            <ShellContent />
        </CockpitProvider>
    );
}

function StartupNoticeBanner({
    notice,
    onDismiss,
}: {
    notice: { tone: "info" | "error"; title: string; details: string[] };
    onDismiss: () => void;
}) {
    return (
        <div style={S.noticeWrap}>
            <PanelState
                tone={notice.tone === "error" ? "error" : "info"}
                compact
                title={notice.title}
                description={notice.details.join(" ")}
            />
            <button type="button" style={S.noticeDismiss} onClick={onDismiss}>
                Fechar
            </button>
        </div>
    );
}

const S: Record<string, React.CSSProperties> = {
    gateRoot: {
        display: "grid",
        alignContent: "start",
        minHeight: "100%",
        paddingTop: 12,
        background: "var(--iccc-bg)",
    },
    odooGateCard: {
        display: "grid",
        gap: 10,
        maxWidth: 420,
    },
    odooGateBtn: {
        justifySelf: "start",
        border: "1px solid var(--iccc-card-border)",
        background: "rgba(255,255,255,0.86)",
        color: "var(--iccc-text)",
        borderRadius: 10,
        padding: "8px 12px",
        fontSize: 12,
        fontWeight: 700,
        cursor: "pointer",
    },
    noticeWrap: {
        display: "grid",
        gap: 8,
        marginBottom: 10,
    },
    noticeDismiss: {
        justifySelf: "end",
        border: "none",
        background: "transparent",
        color: "var(--iccc-text-muted)",
        fontSize: 11,
        fontWeight: 700,
        cursor: "pointer",
        padding: "0 4px",
    },
};
