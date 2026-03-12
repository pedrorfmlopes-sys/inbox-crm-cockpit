import React, { useEffect, useRef, useState } from "react";
import { createPortal } from "react-dom";

export const HelpHint: React.FC<{
    text: string;
    title?: string;
}> = ({ text, title = "Ajuda" }) => {
    const [open, setOpen] = useState(false);
    const [alignRight, setAlignRight] = useState(false);
    const [openUpward, setOpenUpward] = useState(false);
    const [position, setPosition] = useState<{ left: number; top: number }>({ left: 0, top: 0 });
    const wrapRef = useRef<HTMLSpanElement | null>(null);
    const tooltipRef = useRef<HTMLSpanElement | null>(null);

    useEffect(() => {
        if (!open || !wrapRef.current || !tooltipRef.current) return;
        const rect = wrapRef.current.getBoundingClientRect();
        const tooltipRect = tooltipRef.current.getBoundingClientRect();
        const viewportWidth = window.innerWidth || document.documentElement.clientWidth || 0;
        const viewportHeight = window.innerHeight || document.documentElement.clientHeight || 0;
        const spaceRight = viewportWidth - rect.left;
        const spaceLeft = rect.right;
        const spaceBelow = viewportHeight - rect.bottom;
        const spaceAbove = rect.top;
        const nextAlignRight = spaceRight < tooltipRect.width + 12 && spaceLeft > spaceRight;
        const nextOpenUpward = spaceBelow < tooltipRect.height + 12 && spaceAbove > spaceBelow;
        setAlignRight(nextAlignRight);
        setOpenUpward(nextOpenUpward);
        setPosition({
            left: nextAlignRight
                ? Math.max(8, rect.right - tooltipRect.width)
                : Math.min(viewportWidth - tooltipRect.width - 8, rect.left),
            top: nextOpenUpward
                ? Math.max(8, rect.top - tooltipRect.height - 8)
                : Math.min(viewportHeight - tooltipRect.height - 8, rect.bottom + 8),
        });
    }, [open]);

    return (
        <span
            ref={wrapRef}
            style={styles.wrap}
            onMouseEnter={() => setOpen(true)}
            onMouseLeave={() => setOpen(false)}
            onFocus={() => setOpen(true)}
            onBlur={() => setOpen(false)}
        >
            <button
                type="button"
                aria-label={title}
                style={styles.button}
            >
                ?
            </button>
            {open
                ? createPortal(
                    <span
                        ref={tooltipRef}
                        role="tooltip"
                        style={{
                            ...styles.tooltip,
                            left: `${position.left}px`,
                            top: `${position.top}px`,
                            transformOrigin: `${alignRight ? "right" : "left"} ${openUpward ? "bottom" : "top"}`,
                        }}
                    >
                        {text}
                    </span>,
                    document.body
                )
                : null}
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
        position: "fixed",
        zIndex: 9999,
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
