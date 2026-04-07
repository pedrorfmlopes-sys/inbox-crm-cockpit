import React from "react";
import { type ClassificationLayoutMode } from "../types";

export interface SummaryTile {
  key: string;
  title: string;
  value: string;
  description: string;
  onClick: () => void;
}

interface Props {
  tiles: SummaryTile[];
  classificationLayoutMode: ClassificationLayoutMode;
  style?: React.CSSProperties;
}

const ClassificationSummaryTiles: React.FC<Props> = ({ tiles, classificationLayoutMode, style }) => {
  return (
    <div style={{ ...S.classificationSummary, ...style }} data-testid="classification-summary">
      {tiles
        .filter((item) => classificationLayoutMode === "advanced" || item.key !== "references")
        .map((item) => (
          <button key={item.key} data-testid={`summary-tile-${item.key}`} type="button" style={S.classificationTile} onClick={item.onClick}>
            <span style={S.classificationTileLabel}>{item.title}</span>
            <span style={S.classificationTileValue}>{item.value}</span>
            <span style={S.classificationTileMeta}>{item.description}</span>
          </button>
        ))}
      <div style={S.classificationModeHint}>
        {classificationLayoutMode === "normal"
          ? "Modo normal: grupo principal, etiquetas e ticket."
          : "Modo avancado: inclui referencias e opcoes finas."}
      </div>
    </div>
  );
};

export default ClassificationSummaryTiles;

const S: Record<string, React.CSSProperties> = {
  classificationSummary: { minHeight: 0, display: "grid", gap: 8, alignContent: "start", overflowY: "auto", paddingRight: 1 },
  classificationTile: { width: "100%", textAlign: "left", borderRadius: 12, border: "1px solid rgba(148,163,184,0.16)", background: "rgba(255,255,255,0.76)", padding: "10px 12px", display: "grid", gap: 3, cursor: "pointer", transition: "all 150ms ease" },
  classificationTileLabel: { fontSize: 8.5, fontWeight: 700, letterSpacing: "0.09em", textTransform: "uppercase", color: "var(--iccc-muted)" },
  classificationTileValue: { fontSize: 11.25, fontWeight: 600, color: "var(--iccc-text)", lineHeight: 1.25 },
  classificationTileMeta: { fontSize: 9.5, lineHeight: 1.3, color: "var(--iccc-muted)" },
  classificationModeHint: { marginTop: 4, padding: "8px 10px", borderRadius: 10, border: "1px dashed rgba(148,163,184,0.22)", background: "rgba(248,250,252,0.82)", color: "var(--iccc-muted)", fontSize: 9.75, lineHeight: 1.4 },
};
