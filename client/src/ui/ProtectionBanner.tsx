import React from "react";
import * as Icons from "./icons";
import { ProtectedProject } from "../modules/crm/excelProvider";

interface ProtectionBannerProps {
    project: ProtectedProject;
    confidence: number;
    reason?: string;
    onDraftRejection: () => void;
    isDrafting?: boolean;
}

export const ProtectionBanner: React.FC<ProtectionBannerProps> = ({
    project,
    confidence,
    reason,
    onDraftRejection,
    isDrafting
}) => {
    const isHighConflict = confidence > 0.95;
    const color = isHighConflict ? "#991b1b" : "#854d0e";
    const bg = isHighConflict ? "#fee2e2" : "#fef9c3";
    const border = isHighConflict ? "#f87171" : "#facc15";

    return (
        <div style={{
            padding: "10px",
            background: bg,
            border: `1px solid ${border}`,
            borderRadius: "6px",
            marginBottom: "12px",
            display: "flex",
            flexDirection: "column",
            gap: "8px",
            boxShadow: "0 2px 4px rgba(0,0,0,0.05)",
        }}>
            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                <div style={{ color }}>
                    {isHighConflict ? <Icons.AlertCircle size={18} /> : <Icons.AlertTriangle size={18} />}
                </div>
                <div style={{ flex: 1 }}>
                    <div style={{ fontWeight: 800, fontSize: "11px", color, textTransform: "uppercase" }}>
                        ALERTA DE PROTEÇÃO: {isHighConflict ? "CONFLITO CRÍTICO" : "PROJETO PROTEGIDO"}
                    </div>
                    <div style={{ fontSize: "12px", fontWeight: 700, color: "black", marginTop: "2px" }}>
                        {project.projectName}
                    </div>
                </div>
            </div>

            <div style={{ fontSize: "11px", color: "#444" }}>
                <b>Distribuidor:</b> {project.distributor}
                {reason && <div style={{ marginTop: "2px", opacity: 0.8 }}><i>({reason})</i></div>}
            </div>

            <button
                onClick={onDraftRejection}
                disabled={isDrafting}
                style={{
                    background: color,
                    color: "white",
                    border: "none",
                    borderRadius: "4px",
                    padding: "6px 10px",
                    fontSize: "11px",
                    fontWeight: 700,
                    cursor: isDrafting ? "wait" : "pointer",
                    display: "flex",
                    alignItems: "center",
                    justifyContent: "center",
                    gap: "6px",
                }}
            >
                {isDrafting ? <Icons.RefreshCw size={12} className="animate-spin" /> : <Icons.MessageSquare size={12} />}
                Gerar Resposta Diplomática
            </button>
        </div>
    );
};
