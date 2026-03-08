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
                if (st.skinId) applySkin(st.skinId);
            } catch {
                // ignore
            }
        })();
    }, []);

    return <CockpitShell />;
}
