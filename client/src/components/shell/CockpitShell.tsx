import React from "react";
import { CockpitProvider, useCockpit } from "./CockpitProvider";
import { Navigation } from "./Navigation";
import { GlobalHeader } from "./GlobalHeader";

// Modules
import { AiCockpit } from "../../modules/ai/AiCockpit";
import { CrmCockpit } from "../../modules/crm/CrmCockpit";
import { FileCockpit } from "../../modules/files/FileCockpit";
import { SettingsPanel } from "../../ui/SettingsPanel";

function ShellContent() {
    const { tab } = useCockpit();

    return (
        <div style={{
            display: "flex",
            flexDirection: "column",
            height: "100vh",
            background: "var(--iccc-bg)",
            color: "var(--iccc-text)",
            fontFamily: "var(--iccc-font)",
            overflow: "hidden"
        }}>
            <GlobalHeader />

            <main className="flex-1 overflow-y-auto relative" style={{ padding: "12px", paddingBottom: 80 }}>
                {tab === "ai" && <AiCockpit />}
                {tab === "crm" && <CrmCockpit />}
                {tab === "files" && <FileCockpit />}
                {tab === "settings" && <SettingsPanel />}
            </main>

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
