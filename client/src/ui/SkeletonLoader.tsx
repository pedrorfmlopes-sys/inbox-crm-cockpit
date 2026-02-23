import React from "react";

interface SkeletonProps {
    width?: string | number;
    height?: string | number;
    borderRadius?: string | number;
    marginTop?: string | number;
    marginBottom?: string | number;
}

export const Skeleton: React.FC<SkeletonProps> = ({
    width = "100%",
    height = "16px",
    borderRadius = "4px",
    marginTop = 0,
    marginBottom = 0
}) => {
    return (
        <div style={{
            width,
            height,
            borderRadius,
            marginTop,
            marginBottom,
            background: "linear-gradient(90deg, #f0f0f0 25%, #f8f8f8 50%, #f0f0f0 75%)",
            backgroundSize: "200% 100%",
            animation: "skeleton-loading 1.5s infinite linear",
        }} />
    );
};

export const OdooCardSkeleton: React.FC = () => {
    return (
        <div style={{
            padding: "8px",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "6px",
            background: "white",
            marginBottom: "6px",
        }}>
            <div style={{ display: "flex", gap: "6px", marginBottom: "8px" }}>
                <Skeleton width="40px" height="12px" />
                <Skeleton width="30px" height="12px" />
            </div>
            <Skeleton width="80%" height="14px" marginBottom="8px" />
            <div style={{ display: "flex", justifyContent: "flex-end" }}>
                <Skeleton width="50px" height="12px" />
            </div>
        </div>
    );
};

// Add global styles for animation if not present
if (typeof document !== "undefined") {
    const styleId = "iccc-skeleton-styles";
    if (!document.getElementById(styleId)) {
        const style = document.createElement("style");
        style.id = styleId;
        style.innerHTML = `
            @keyframes skeleton-loading {
                0% { background-position: 200% 0; }
                100% { background-position: -200% 0; }
            }
        `;
        document.head.appendChild(style);
    }
}
