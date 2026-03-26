import React, { useEffect, useMemo } from "react";
import { CockpitProvider } from "@/components/shell/CockpitProvider";
import { GroupManagerCockpit } from "@/modules/crm/GroupManagerCockpit";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import "../../global.css";

function getInitialSection(): "settings" | "labels" | "tickets" {
  const params = new URLSearchParams(window.location.search);
  const section = String(params.get("section") || "settings").trim().toLowerCase();
  if (section === "labels" || section === "tickets") return section;
  return "settings";
}

export default function GroupSettingsApp(): JSX.Element {
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
          background: "var(--iccc-bg)",
          color: "var(--iccc-text)",
          fontFamily: "var(--iccc-font)",
          overflow: "hidden",
        }}
      >
        <GroupManagerCockpit initialView={initialSection} standaloneSettings />
      </div>
    </CockpitProvider>
  );
}
