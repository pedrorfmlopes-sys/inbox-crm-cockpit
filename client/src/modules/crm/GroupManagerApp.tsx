import React, { useEffect } from "react";
import { CockpitProvider } from "@/components/shell/CockpitProvider";
import { GroupManagerCockpit } from "@/modules/crm/GroupManagerCockpit";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import "../../global.css";

export default function GroupManagerApp(): JSX.Element {
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
          minHeight: "100vh",
          padding: 12,
          boxSizing: "border-box",
          background: "var(--iccc-bg)",
          color: "var(--iccc-text)",
          fontFamily: "var(--iccc-font)",
        }}
      >
        <GroupManagerCockpit />
      </div>
    </CockpitProvider>
  );
}
