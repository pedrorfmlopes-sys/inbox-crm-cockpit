import React from "react";
import { CockpitProvider, useCockpit } from "./CockpitProvider";
import { Navigation } from "./Navigation";
import { GlobalHeader } from "./GlobalHeader";

// Modules
import { AiCockpit } from "../../modules/ai/AiCockpit";
import { CrmCockpit } from "../../modules/crm/CrmCockpit";
import { FileCockpit } from "../../modules/files/FileCockpit";
import { SettingsPanel } from "../../ui/SettingsPanel";

import { LoginCockpit } from "../../modules/auth/LoginCockpit";

function ShellContent() {
    const cockpit = useCockpit();
    if (!cockpit) return null;
    const { tab, isAuthenticated, isLoading } = cockpit as any;

    if (isLoading) {
        return (
            <div style={{ display: "flex", justifyContent: "center", alignItems: "center", height: "100%" }}>
                <div style={{ color: "var(--iccc-text-muted)" }}>A carregar...</div>
            </div>
        );
    }

    if (!isAuthenticated) {
        return <LoginCockpit />;
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
                {tab === "ai" && <AiCockpit />}
                {tab === "crm" && <CrmCockpit />}
                {tab === "files" && <FileCockpit />}
                {tab === "settings" && <SettingsPanel />}
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
