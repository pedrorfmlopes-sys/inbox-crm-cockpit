import React from "react";
import { type RelatedEmailEntry } from "@/api";
import { type ClassificationFocus, type ApplyDialogScopeMode } from "../types";
import { makeEmailKey, buildCompactEmailMeta, buildEmailPreviewText } from "../documentUtils";

export interface ApplyDialogProps {
  isOpen: boolean;
  onClose: () => void;
  section: ClassificationFocus;
  scopeMode: ApplyDialogScopeMode;
  setScopeMode: (mode: ApplyDialogScopeMode) => void;
  currentScopeEmail: RelatedEmailEntry | null;
  caseScopeEmails: RelatedEmailEntry[];
  selectedEmailKeys: string[];
  setSelectedEmailKeys: (keys: string[]) => void;
  expandedEmailKeys: string[];
  toggleExpandedEmailKey: (key: string) => void;
  toggleEmailKey: (key: string) => void;
  status: string;
  actionBusy: boolean;
  handleConfirm: () => void;
}

const ApplyDialog: React.FC<ApplyDialogProps> = ({
  isOpen,
  onClose,
  section,
  scopeMode,
  setScopeMode,
  currentScopeEmail,
  caseScopeEmails,
  selectedEmailKeys,
  setSelectedEmailKeys,
  expandedEmailKeys,
  toggleExpandedEmailKey,
  toggleEmailKey,
  status,
  actionBusy,
  handleConfirm,
}) => {
  if (!isOpen) return null;

  const sectionLabel = section === "principal"
    ? "Grupo principal"
    : section === "labels"
      ? "Etiquetas"
      : section === "ticket"
        ? "Ticket"
        : section === "references"
          ? "Referencias"
          : "Classificacao";

  const displayEmails = scopeMode === "current"
    ? (currentScopeEmail ? [currentScopeEmail] : [])
    : caseScopeEmails;

  const manualSelectionEnabled = scopeMode === "selected";
  const showEmailList = scopeMode !== "current";

  return (
    <div style={S.modalBackdrop} data-testid="apply-dialog">
      <div style={S.modalSheet}>
        <div style={S.modalHeader}>
          <div>
            <div style={S.kicker}>Aplicar alteracoes</div>
            <div style={S.modalTitle}>{sectionLabel}</div>
          </div>
          <button data-testid="apply-dialog-cancel" type="button" style={S.secondaryBtn} onClick={onClose} disabled={actionBusy}>Cancelar</button>
        </div>

        <div style={S.modalScopeRow}>
          <button type="button" style={scopeMode === "current" ? S.scopeChipOn : S.scopeChip} onClick={() => setScopeMode("current")} disabled={actionBusy}>So este email</button>
          <button type="button" style={scopeMode === "selected" ? S.scopeChipOn : S.scopeChip} onClick={() => setScopeMode("selected")} disabled={actionBusy}>Emails selecionados</button>
          <button type="button" style={scopeMode === "case_all" ? S.scopeChipOn : S.scopeChip} onClick={() => setScopeMode("case_all")} disabled={actionBusy}>Todos os emails do caso</button>
        </div>

        <div style={S.modalBlock}>
          <div style={S.modalBlockHeader}>
            <div style={S.editorBlockTitle}>{scopeMode === "current" ? "Email alvo" : "Escolher emails"}</div>
            {manualSelectionEnabled ? (
              <button 
                type="button" 
                style={S.linkBtn} 
                onClick={() => setSelectedEmailKeys(caseScopeEmails.map((email) => makeEmailKey(email)).filter(Boolean) as string[])} 
                disabled={actionBusy}
              >
                Selecionar todos
              </button>
            ) : null}
          </div>
          {!showEmailList && currentScopeEmail ? (
            <div style={S.applySingleEmailSummary}>
              <div style={S.applyEmailSummaryHead}>
                <div style={S.applyEmailSummaryTitle}>{currentScopeEmail.subject || "(sem assunto)"}</div>
                <button type="button" style={S.chevronBtn} onClick={() => toggleExpandedEmailKey(makeEmailKey(currentScopeEmail))}>
                  {expandedEmailKeys.includes(makeEmailKey(currentScopeEmail)) ? "\u2303" : "\u2304"}
                </button>
              </div>
              <div style={S.applyEmailSummaryMeta}>{buildCompactEmailMeta(currentScopeEmail)}</div>
              {expandedEmailKeys.includes(makeEmailKey(currentScopeEmail)) ? (
                <div style={S.applyEmailPreview}>{buildEmailPreviewText(currentScopeEmail) || "Sem preview resumido para este email."}</div>
              ) : null}
            </div>
          ) : (
            <div style={S.applyEmailList}>
              {displayEmails.map((email) => {
                const emailKey = makeEmailKey(email);
                const isSelected = selectedEmailKeys.includes(emailKey);
                const isExpanded = expandedEmailKeys.includes(emailKey);
                return (
                  <div key={emailKey} style={isSelected ? S.applyEmailRowOn : S.applyEmailRow}>
                    <div style={S.applyEmailRowTop}>
                      <div style={S.applyEmailMain}>
                        {manualSelectionEnabled ? (
                          <input type="checkbox" checked={isSelected} onChange={() => toggleEmailKey(emailKey)} disabled={actionBusy} />
                        ) : (
                          <span style={isSelected ? S.applyScopeBadgeOn : S.applyScopeBadge}>{isSelected ? "Incluido" : "Omitir"}</span>
                        )}
                        <span style={S.applyEmailSubject}>{email.subject || "(sem assunto)"}</span>
                      </div>
                      <div style={S.applyEmailRowTail}>
                        <span style={S.applyEmailMeta}>{buildCompactEmailMeta(email)}</span>
                        <button type="button" style={S.chevronBtn} onClick={() => toggleExpandedEmailKey(emailKey)}>
                          {isExpanded ? "\u2303" : "\u2304"}
                        </button>
                      </div>
                    </div>
                    {isExpanded ? (
                      <div style={S.applyEmailPreview}>{buildEmailPreviewText(email) || "Sem preview resumido para este email."}</div>
                    ) : null}
                  </div>
                );
              })}
            </div>
          )}
        </div>

        <div style={S.modalFooter}>
          <div style={{ flex: 1, display: "flex", alignItems: "center", gap: 10 }}>
            {status ? <div style={{ fontSize: 11, fontWeight: 600, color: "var(--iccc-muted)" }}>{status}</div> : null}
            {actionBusy ? <span style={{ fontSize: 10, color: "#1d4ed8", fontWeight: 700 }}>A processar...</span> : null}
          </div>
          <button data-testid="apply-dialog-cancel" type="button" style={S.secondaryBtn} onClick={onClose} disabled={actionBusy}>Cancelar</button>
          <button data-testid="apply-dialog-confirm" type="button" style={S.primaryBtn} onClick={handleConfirm} disabled={actionBusy || (manualSelectionEnabled && !selectedEmailKeys.length)}>
            {actionBusy ? "A aplicar..." : "Confirmar e aplicar"}
          </button>
        </div>
      </div>
    </div>
  );
};

const S: Record<string, React.CSSProperties> = {
  modalBackdrop: { position: "fixed", top: 0, left: 0, right: 0, bottom: 0, background: "rgba(0,0,0,0.5)", display: "flex", alignItems: "center", justifyContent: "center", zIndex: 1000, padding: 20 },
  modalSheet: { background: "var(--skin-bg-main)", borderRadius: 12, width: "100%", maxWidth: 500, maxHeight: "90vh", display: "flex", flexDirection: "column", boxShadow: "0 20px 25px -5px rgba(0,0,0,0.1), 0 10px 10px -5px rgba(0,0,0,0.04)" },
  modalHeader: { padding: "16px 20px", borderBottom: "1px solid var(--skin-border-main)", display: "flex", justifyContent: "space-between", alignItems: "flex-start" },
  kicker: { fontSize: 10, fontWeight: 700, color: "var(--skin-accent-main)", textTransform: "uppercase", letterSpacing: "0.05em", marginBottom: 2 },
  modalTitle: { fontSize: 18, fontWeight: 700, color: "var(--skin-text-main)" },
  modalScopeRow: { padding: "12px 20px", display: "flex", gap: 8, background: "var(--skin-bg-muted)", borderBottom: "1px solid var(--skin-border-main)" },
  scopeChip: { flex: 1, padding: "8px 4px", fontSize: 11, borderRadius: 6, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-main)", color: "var(--skin-text-muted)", cursor: "pointer", fontWeight: 500 },
  scopeChipOn: { flex: 1, padding: "8px 4px", fontSize: 11, borderRadius: 6, border: "1px solid var(--skin-accent-main)", background: "var(--skin-bg-active)", color: "var(--skin-accent-main)", cursor: "pointer", fontWeight: 600 },
  modalBlock: { padding: "16px 20px", flex: 1, overflowY: "auto", display: "flex", flexDirection: "column", gap: 12 },
  modalBlockHeader: { display: "flex", justifyContent: "space-between", alignItems: "center" },
  editorBlockTitle: { fontSize: 11, fontWeight: 700, color: "var(--skin-text-muted)", textTransform: "uppercase" },
  linkBtn: { background: "none", border: "none", padding: 0, color: "var(--skin-accent-main)", fontSize: 11, cursor: "pointer", textDecoration: "underline" },
  applySingleEmailSummary: { padding: 12, borderRadius: 8, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-muted)" },
  applyEmailSummaryHead: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 },
  applyEmailSummaryTitle: { fontSize: 13, fontWeight: 600, color: "var(--skin-text-main)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  applyEmailSummaryMeta: { fontSize: 11, color: "var(--skin-text-muted)", marginTop: 2 },
  applyEmailPreview: { fontSize: 11, color: "var(--skin-text-muted)", background: "var(--skin-bg-main)", padding: 10, borderRadius: 6, border: "1px solid var(--skin-border-main)", marginTop: 8, lineHeight: "1.4" },
  applyEmailList: { display: "flex", flexDirection: "column", gap: 6 },
  applyEmailRow: { padding: "8px 12px", borderRadius: 8, border: "1px solid var(--skin-border-main)", background: "var(--skin-bg-main)" },
  applyEmailRowOn: { padding: "8px 12px", borderRadius: 8, border: "1px solid var(--skin-accent-main)", background: "var(--skin-bg-active)" },
  applyEmailRowTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 10 },
  applyEmailMain: { display: "flex", alignItems: "center", gap: 10, flex: 1, overflow: "hidden" },
  applyScopeBadge: { fontSize: 9, fontWeight: 700, padding: "2px 6px", borderRadius: 4, background: "var(--skin-bg-muted)", color: "var(--skin-text-muted)" },
  applyScopeBadgeOn: { fontSize: 9, fontWeight: 700, padding: "2px 6px", borderRadius: 4, background: "var(--skin-bg-success-muted)", color: "var(--skin-bg-success)" },
  applyEmailSubject: { fontSize: 12, fontWeight: 500, color: "var(--skin-text-main)", overflow: "hidden", textOverflow: "ellipsis", whiteSpace: "nowrap" },
  applyEmailRowTail: { display: "flex", alignItems: "center", gap: 8, flexShrink: 0 },
  applyEmailMeta: { fontSize: 11, color: "var(--skin-text-muted)" },
  chevronBtn: { background: "none", border: "none", width: 24, height: 24, display: "flex", alignItems: "center", justifyContent: "center", cursor: "pointer", color: "var(--skin-text-muted)", fontSize: 14 },
  modalFooter: { padding: "16px 20px", borderTop: "1px solid var(--skin-border-main)", display: "flex", gap: 10, alignItems: "center", background: "var(--skin-bg-muted)", borderBottomLeftRadius: 12, borderBottomRightRadius: 12 },
  primaryBtn: { padding: "10px 16px", fontSize: 13, fontWeight: 600, borderRadius: 8, background: "var(--skin-accent-main)", color: "white", border: "none", cursor: "pointer" },
  secondaryBtn: { padding: "10px 16px", fontSize: 13, fontWeight: 600, borderRadius: 8, background: "var(--skin-bg-main)", color: "var(--skin-text-main)", border: "1px solid var(--skin-border-main)", cursor: "pointer" },
};

export default ApplyDialog;
