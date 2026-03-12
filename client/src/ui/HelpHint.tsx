import React, { useState } from "react";

export const HelpHint: React.FC<{
    text: string;
    title?: string;
}> = ({ text, title = "Ajuda" }) => {
    const [open, setOpen] = useState(false);

    return (
        <span
            style={styles.wrap}
            onMouseEnter={() => setOpen(true)}
            onMouseLeave={() => setOpen(false)}
            onFocus={() => setOpen(true)}
            onBlur={() => setOpen(false)}
        >
            <button
                type="button"
                aria-label={title}
                title={title}
                style={styles.button}
            >
                ?
            </button>
            {open ? (
                <span role="tooltip" style={styles.tooltip}>
                    {text}
                </span>
            ) : null}
        </span>
    );
};

const styles: Record<string, React.CSSProperties> = {
    wrap: {
        position: "relative",
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        flexShrink: 0,
    },
    button: {
        width: "14px",
        height: "14px",
        minWidth: "14px",
        borderRadius: "999px",
        border: "1px solid #C1C7D0",
        background: "#FFFFFF",
        color: "#6B778C",
        fontSize: "9px",
        fontWeight: 800,
        lineHeight: 1,
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        padding: 0,
        cursor: "help",
    },
    tooltip: {
        position: "absolute",
        top: "18px",
        right: 0,
        zIndex: 30,
        minWidth: "170px",
        maxWidth: "220px",
        padding: "6px 8px",
        borderRadius: "8px",
        border: "1px solid #C1C7D0",
        background: "#172B4D",
        color: "#FFFFFF",
        fontSize: "10px",
        lineHeight: 1.35,
        boxShadow: "0 8px 18px rgba(9, 30, 66, 0.18)",
        pointerEvents: "none",
        textAlign: "left",
        whiteSpace: "normal",
    },
};
