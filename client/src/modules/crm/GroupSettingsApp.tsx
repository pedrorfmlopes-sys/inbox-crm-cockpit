import React, { Suspense, useEffect, useMemo, useState } from "react";
import { CockpitProvider } from "@/components/shell/CockpitProvider";
import { requestCockpitHostAction } from "@/office";
import { buildGroupsSettingsPatch } from "@/modules/crm/groups-v1/settings/groupsModuleSettings";
import type { GroupsSettingsSection } from "@/modules/crm/groups-v1/settings/GroupsSettingsPanel";
import { getSettings, saveSettings, type CockpitSettingsV1 } from "@/settings";
import { applySkin } from "@/ui/skins";
import "../../global.css";

const GroupManagerCockpitLazy = React.lazy(async () => {
  const module = await import("@/modules/crm/GroupManagerCockpit");
  return { default: module.GroupManagerCockpit };
});

const GroupsSettingsPanelLazy = React.lazy(async () => {
  const module = await import("@/modules/crm/groups-v1/settings/GroupsSettingsPanel");
  return { default: module.GroupsSettingsPanel };
});

type GroupSettingsSurface = "manager" | "groups-tab";

function getQueryParams() {
  const params = new URLSearchParams(window.location.search);
  return {
    surface: String(params.get("surface") || "").trim().toLowerCase(),
    section: String(params.get("section") || "").trim().toLowerCase(),
  };
}

function getInitialSurface(): GroupSettingsSurface {
  return getQueryParams().surface === "groups-tab" ? "groups-tab" : "manager";
}

function getInitialManagerSection(): "settings" | "labels" | "tickets" {
  const { section } = getQueryParams();
  if (section === "labels" || section === "tickets") return section;
  return "settings";
}

function getInitialGroupsTabSection(): GroupsSettingsSection {
  const { section } = getQueryParams();
  if (
    section === "general"
    || section === "intermediate_storage"
    || section === "attachments"
    || section === "cleanup"
    || section === "warnings"
    || section === "migration"
    || section === "maintenance"
    || section === "explore"
    || section === "about"
  ) {
    return section;
  }
  return "general";
}

function SettingsShell({ children }: { children: React.ReactNode }) {
  return (
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
      {children}
    </div>
  );
}

function LoadingState() {
  return (
    <SettingsShell>
      <div
        style={{
          display: "grid",
          placeItems: "center",
          borderRadius: 18,
          border: "1px solid rgba(148,163,184,0.18)",
          background: "rgba(255,255,255,0.82)",
          fontSize: 13,
          fontWeight: 600,
          color: "#334155",
        }}
      >
        A preparar os settings...
      </div>
    </SettingsShell>
  );
}

export default function GroupSettingsApp(): JSX.Element {
  const surface = useMemo(() => getInitialSurface(), []);
  const initialManagerSection = useMemo(() => getInitialManagerSection(), []);
  const initialGroupsTabSection = useMemo(() => getInitialGroupsTabSection(), []);
  const [settings, setSettings] = useState<CockpitSettingsV1 | null>(null);
  const [loading, setLoading] = useState(true);
  const [status, setStatus] = useState<{ tone: "success" | "error"; text: string } | null>(null);

  useEffect(() => {
    let alive = true;
    void (async () => {
      try {
        const nextSettings = await getSettings();
        if (!alive) return;
        setSettings(nextSettings);
        if (nextSettings.skinId) applySkin(nextSettings.skinId);
      } catch (error) {
        if (!alive) return;
        setStatus({
          tone: "error",
          text: error instanceof Error ? error.message : "Nao foi possivel carregar os settings.",
        });
      } finally {
        if (alive) setLoading(false);
      }
    })();
    return () => {
      alive = false;
    };
  }, []);

  async function handleClose() {
    const closed = await requestCockpitHostAction({ type: "close" });
    if (closed) return;

    if (window.opener && window.opener !== window) {
      try {
        window.close();
        if (window.closed) return;
      } catch {
        // continue to browser fallback navigation
      }
    }

    const fallbackUrl = new URL(window.location.href);
    fallbackUrl.searchParams.delete("view");
    fallbackUrl.searchParams.delete("surface");
    fallbackUrl.searchParams.delete("section");
    window.location.assign(fallbackUrl.toString());
  }

  async function handleSaveGroupsTabSettings(nextTabSettings: NonNullable<CockpitSettingsV1["groups"]["tab"]>) {
    try {
      const saved = await saveSettings(buildGroupsSettingsPatch(settings, { tab: nextTabSettings }));
      setSettings(saved);
      setStatus({ tone: "success", text: "Settings da aba Groups guardados." });
    } catch (error) {
      const message = error instanceof Error ? error.message : "Nao foi possivel guardar os settings da aba Groups.";
      setStatus({ tone: "error", text: message });
      throw error;
    }
  }

  if (loading) {
    return (
      <CockpitProvider>
        <LoadingState />
      </CockpitProvider>
    );
  }

  return (
    <CockpitProvider>
      <SettingsShell>
        <Suspense fallback={<LoadingState />}>
          {surface === "groups-tab" ? (
            <GroupsSettingsPanelLazy
              open={true}
              value={settings?.groups?.tab || null}
              onClose={handleClose}
              onSave={handleSaveGroupsTabSettings}
              initialSection={initialGroupsTabSection}
              statusMessage={status?.text || ""}
              statusTone={status?.tone || "success"}
            />
          ) : (
            <GroupManagerCockpitLazy initialView={initialManagerSection} standaloneSettings />
          )}
        </Suspense>
      </SettingsShell>
    </CockpitProvider>
  );
}
