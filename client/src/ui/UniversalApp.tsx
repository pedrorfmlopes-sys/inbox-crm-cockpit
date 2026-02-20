import React, { useEffect } from "react";
import { CockpitShell } from "@/components/shell/CockpitShell";
import { getSettings } from "@/settings";
import { applySkin } from "@/ui/skins";
import "../global.css";

export default function UniversalApp() {
    useEffect(() => {
        (async () => {
            try {
                const st = await getSettings();
                applySkin(st.skinId || "vibrant"); // Default to vibrant for the new experience
            } catch {
                applySkin("vibrant");
            }
        })();
    }, []);

    return <CockpitShell />;
}
