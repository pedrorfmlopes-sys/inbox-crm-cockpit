import React from "react";
import { UNIFIED_STATUS_LEGEND } from "@/statusUtils";

const StatusLegend: React.FC = () => {
  return (
    <div style={S.legend}>
      {UNIFIED_STATUS_LEGEND.map((status) => (
        <div key={status.key} style={S.legendItem}>
          <div style={{ ...S.legendDot, background: status.hex }} />
          <span>{status.label}</span>
        </div>
      ))}
    </div>
  );
};

const S: Record<string, React.CSSProperties> = {
  legend: { display: "flex", gap: 12, flexWrap: "wrap", marginTop: 8, padding: "8px 0", borderTop: "1px solid var(--skin-border-main)" },
  legendItem: { display: "flex", alignItems: "center", gap: 6, fontSize: 10, color: "var(--skin-text-muted)" },
  legendDot: { width: 8, height: 8, borderRadius: "50%" },
};

export default StatusLegend;
