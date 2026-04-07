import React from "react";
import * as Icons from "@/ui/icons";
import { type ClassificationFocus, type ClassificationLayoutMode } from "../types";

interface Props {
  classificationFocus: ClassificationFocus;
  classificationLayoutMode: ClassificationLayoutMode;
  onBack: () => void;
  onApply: (focus: ClassificationFocus) => void;
  canApply: boolean;
  actionBusy: boolean;
}

const ClassificationEditorHeader: React.FC<Props> = ({
  classificationFocus,
  classificationLayoutMode,
  onBack,
  onApply,
  canApply,
  actionBusy,
}) => {
  const focusTitle = classificationFocus === "principal"
    ? "Grupo principal"
    : classificationFocus === "labels"
      ? "Etiquetas"
      : classificationFocus === "ticket"
        ? "Ticket"
        : "Referencias";

  return (
    <div style={S.editorHeader}>
      <div style={S.editorHeaderMeta}>
        <div style={S.sectionTitle}>Classificacao</div>
        <div style={S.editorHeaderTitle}>{focusTitle}</div>
        <div style={S.editorModeText}>{classificationLayoutMode === "advanced" ? "Modo avancado" : "Modo normal"}</div>
      </div>
      <div style={S.editorHeaderActions}>
        <button type="button" style={S.secondaryBtn} onClick={onBack}>Voltar</button>
        <button type="button" style={S.primaryBtn} onClick={() => onApply(classificationFocus)} disabled={actionBusy || !canApply}>
          <Icons.Save size={12} />
          Aplicar
        </button>
      </div>
    </div>
  );
};

export default ClassificationEditorHeader;

const S: Record<string, React.CSSProperties> = {
  editorHeader: { display: "flex", alignItems: "flex-start", justifyContent: "space-between", gap: 10, flexWrap: "wrap", borderBottom: "1px solid rgba(148,163,184,0.14)", paddingBottom: 10, marginBottom: 4 },
  editorHeaderMeta: { display: "grid", gap: 3 },
  sectionTitle: { fontSize: 9.5, fontWeight: 800, textTransform: "uppercase", letterSpacing: "0.1em", color: "rgba(15,23,42,0.82)" },
  editorHeaderTitle: { fontSize: 13.5, fontWeight: 650, color: "var(--iccc-text)" },
  editorModeText: { fontSize: 10, color: "var(--iccc-muted)" },
  editorHeaderActions: { display: "flex", alignItems: "center", gap: 8, flexWrap: "wrap" },
  primaryBtn: { height: 30, padding: "0 11px", borderRadius: 999, border: "1px solid rgba(37,99,235,0.18)", background: "linear-gradient(180deg,#3b82f6 0%, #2563eb 100%)", color: "#fff", fontSize: 10.5, fontWeight: 700, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer", boxShadow: "0 4px 10px rgba(37,99,235,0.14)" },
  secondaryBtn: { height: 28, padding: "0 10px", borderRadius: 999, border: "1px solid rgba(148,163,184,0.18)", background: "rgba(255,255,255,0.9)", color: "var(--iccc-text)", fontSize: 10.5, fontWeight: 600, display: "inline-flex", alignItems: "center", gap: 6, cursor: "pointer" },
};
