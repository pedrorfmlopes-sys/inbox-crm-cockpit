import React from "react";

export type PanelStateTone = "loading" | "empty" | "error" | "success" | "info";

export function PanelState({
    tone,
    title,
    description,
    compact = false,
}: {
    tone: PanelStateTone;
    title: string;
    description?: string;
    compact?: boolean;
}): JSX.Element {
    const palette = getPalette(tone);

    return (
        <div
            style={{
                border: `1px solid ${palette.border}`,
                background: palette.background,
                color: palette.title,
                borderRadius: compact ? 8 : 12,
                padding: compact ? "10px 12px" : "14px 16px",
                display: "grid",
                gap: 4,
            }}
        >
            <div style={{ fontSize: 12, fontWeight: 700 }}>{title}</div>
            {description ? (
                <div style={{ fontSize: 11, lineHeight: 1.4, color: palette.description }}>
                    {description}
                </div>
            ) : null}
        </div>
    );
}

function getPalette(tone: PanelStateTone) {
    if (tone === "loading") {
        return { background: "#F4F5F7", border: "#DFE1E6", title: "#172B4D", description: "#42526E" };
    }
    if (tone === "empty") {
        return { background: "#FAFBFC", border: "#DFE1E6", title: "#42526E", description: "#6B778C" };
    }
    if (tone === "error") {
        return { background: "#FFEBE6", border: "#FFBDAD", title: "#BF2600", description: "#BF2600" };
    }
    if (tone === "success") {
        return { background: "#E3FCEF", border: "#ABF5D1", title: "#006644", description: "#006644" };
    }
    return { background: "#DEEBFF", border: "#B3D4FF", title: "#0747A6", description: "#0747A6" };
}
