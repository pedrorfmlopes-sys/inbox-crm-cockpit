import React, { useEffect, useMemo } from "react";
import { CockpitProvider, useCockpit, type SettingsPanelSection } from "@/components/shell/CockpitProvider";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import { SettingsPanel } from "@/ui/SettingsPanel";
import "../global.css";

function isSettingsPanelSection(value: string | null): value is SettingsPanelSection {
  return value === "general"
    || value === "conns"
    || value === "ai"
    || value === "persona"
    || value === "signature"
    || value === "references"
    || value === "groups"
    || value === "crm2layout"
    || value === "protection";
}

function getInitialSection(): SettingsPanelSection {
  const params = new URLSearchParams(window.location.search);
  const section = params.get("section");
  return isSettingsPanelSection(section) ? section : "general";
}

function SettingsAppBody({ initialSection }: { initialSection: SettingsPanelSection }) {
  const { setSettingsSection } = useCockpit();

  useEffect(() => {
    setSettingsSection(initialSection);
  }, [initialSection, setSettingsSection]);

  return (
    <div
      style={{
        height: "100%",
        overflowY: "auto",
      }}
    >
      <SettingsPanel />
    </div>
  );
}

export default function SettingsApp(): JSX.Element {
  const initialSection = useMemo(() => getInitialSection(), []);

  useEffect(() => {
    (async () => {
      try {
        const settings = await getSettings();
        if (settings.skinId) applySkin(settings.skinId);
      } catch {
        // ignore
      }
    })();
  }, []);

  return (
    <CockpitProvider>
      <div
        style={{
          height: "100vh",
          padding: 12,
          boxSizing: "border-box",
          display: "grid",
          background: "var(--iccc-bg)",
          color: "var(--iccc-text)",
          fontFamily: "var(--iccc-font)",
          overflow: "hidden",
        }}
      >
        <SettingsAppBody initialSection={initialSection} />
      </div>
    </CockpitProvider>
  );
}
