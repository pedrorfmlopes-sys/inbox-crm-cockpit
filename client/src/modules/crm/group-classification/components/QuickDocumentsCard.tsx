import { type RelatedEmailEntry } from "@/api";
import React from "react";
import { 
  makeAttachmentKey, 
  getStudioAttachmentRemoteId, 
  isStudioAttachmentHydrated, 
  isStudioAttachmentHydratedInCollection 
} from "../documentUtils";

export interface QuickDocumentsCardProps {
  style: React.CSSProperties;
  quickDocumentAttachments: any[];
  selectedAttachmentPreviewKey: string;
  previewMode: "email" | "document" | "reply" | "forward";
  expandedQuickDocumentKeys: string[];
  quickDocumentHiddenCount: number;
  showHiddenQuickDocuments: boolean;
  setShowHiddenQuickDocuments: (value: boolean) => void;
  handleOpenQuickAttachment: (email: RelatedEmailEntry, attachment: any) => void;
  handleSetQuickAttachmentHidden: (email: RelatedEmailEntry, attachment: any, hidden: boolean) => void;
  toggleExpandedQuickDocumentKey: (key: string) => void;
  actionBusy: boolean;
}

const QuickDocumentsCard: React.FC<QuickDocumentsCardProps> = ({
  style,
  quickDocumentAttachments,
  selectedAttachmentPreviewKey,
  previewMode,
  expandedQuickDocumentKeys,
  quickDocumentHiddenCount,
  showHiddenQuickDocuments,
  setShowHiddenQuickDocuments,
  handleOpenQuickAttachment,
  handleSetQuickAttachmentHidden,
  toggleExpandedQuickDocumentKey,
  actionBusy
}) => {
  return (
    <section style={style} data-testid="quick-documents-card">
      <div style={S.sectionHeaderCompact}>
        <div>
          <div style={S.sectionTitle}>Documentos Rapidos</div>
          <div style={S.sectionSubtitle}>Anexos detetados no contexto deste caso</div>
        </div>
        {quickDocumentHiddenCount > 0 ? (
          <button 
            type="button" 
            style={S.linkBtn} 
            onClick={() => setShowHiddenQuickDocuments(!showHiddenQuickDocuments)}
          >
            {showHiddenQuickDocuments ? "Esconder silenciados" : `Ver ${quickDocumentHiddenCount} silenciados`}
          </button>
        ) : null}
      </div>
      <div style={S.topCardScroll}>
        {!quickDocumentAttachments.length ? (
          <div style={S.panelState}>
            <div style={S.panelStateTitle}>Nenhum documento</div>
            <div style={S.panelStateDesc}>Nao foram encontrados anexos relevantes nestes emails.</div>
          </div>
        ) : null}
        {quickDocumentAttachments.map((entry, idx) => {
          if (!entry) return null;
          // Handle both {email, attachment} structure and direct attachment structures safely
          const email = entry.email;
          const attachment = entry.attachment || (entry.key || entry.id ? entry : null);
          
          if (!attachment) return null;
          
          const attachmentKey = String(entry.scopedKey || makeAttachmentKey(attachment) || "").trim();
          const remoteId = getStudioAttachmentRemoteId(attachment);
          if (!attachmentKey) return null;
          
          // Guard against missing email when checking hydration in collection
          const emailAttachments = Array.isArray(email?.attachments) ? email.attachments : [];
          const hydrated = isStudioAttachmentHydrated(attachment) || 
                          isStudioAttachmentHydratedInCollection(emailAttachments, attachmentKey);
          
          const active = attachmentKey === selectedAttachmentPreviewKey && previewMode === "document";
          const expanded = expandedQuickDocumentKeys.includes(attachmentKey);
          
          return (
            <div 
              key={`quick-doc-${attachmentKey}-${idx}`} 
              style={active ? S.quickDocOn : S.quickDoc}
              onClick={() => email && handleOpenQuickAttachment(email, attachment)}
            >
              <div style={S.quickDocTop}>
                <div style={S.quickDocName}>{attachment.name || "Sem nome"}</div>
                <div style={S.quickDocTools}>
                  {hydrated ? <span style={S.pillOk}>OK</span> : <span style={S.pillWait}>CLOUD</span>}
                  <button 
                    type="button" 
                    style={S.toolBtn} 
                    disabled={actionBusy} 
                    onClick={(event) => {
                      event.stopPropagation();
                      if (email) handleSetQuickAttachmentHidden(email, attachment, !attachment.isHidden);
                    }}
                  >
                    {attachment.isHidden ? "Restaurar" : "Silenciar"}
                  </button>
                  <button 
                    type="button" 
                    style={S.chevronBtn} 
                    onClick={(event) => {
                      event.stopPropagation();
                      toggleExpandedQuickDocumentKey(attachmentKey);
                    }}
                  >
                    {expanded ? "\u2303" : "\u2304"}
                  </button>
                </div>
              </div>
              {expanded ? (
                <div style={S.quickDocSnippet}>
                  ID Office: {attachment.id || "--"}
                  <br />
                  Chave: {attachmentKey}
                  {remoteId ? <><br />ID Remoto: {remoteId}</> : null}
                  {attachment.contentType ? <><br />Tipo: {attachment.contentType}</> : null}
                </div>
              ) : null}
            </div>
          );
        })}
      </div>
    </section>
  );
};

const S: Record<string, React.CSSProperties> = {
  sectionHeaderCompact: { display: "flex", justifyContent: "space-between", alignItems: "flex-start", marginBottom: 8 },
  sectionTitle: { fontSize: 12, fontWeight: 600, color: "var(--skin-text-main)" },
  sectionSubtitle: { fontSize: 10, color: "var(--skin-text-muted)" },
  linkBtn: { background: "none", border: "none", padding: 0, color: "var(--skin-accent-main)", fontSize: 10, cursor: "pointer", textDecoration: "underline" },
  topCardScroll: { flex: 1, overflowY: "auto", minHeight: 0, display: "flex", flexDirection: "column", gap: 1 },
  panelState: { padding: "12px 0", textAlign: "center" },
  panelStateTitle: { fontSize: 11, fontWeight: 500, color: "var(--skin-text-main)" },
  panelStateDesc: { fontSize: 10, color: "var(--skin-text-muted)" },
  quickDoc: { padding: "6px 8px", background: "var(--skin-bg-card)", border: "1px solid var(--skin-border-main)", borderRadius: 4, cursor: "pointer", display: "flex", flexDirection: "column", gap: 2, textAlign: "left", width: "100%", position: "relative" },
  quickDocOn: { padding: "6px 8px", background: "var(--skin-bg-active)", border: "1px solid var(--skin-accent-main)", borderRadius: 4, cursor: "pointer", display: "flex", flexDirection: "column", gap: 2, textAlign: "left", width: "100%", position: "relative" },
  quickDocTop: { display: "flex", justifyContent: "space-between", alignItems: "center", gap: 8 },
  quickDocName: { fontSize: 11, fontWeight: 500, color: "var(--skin-text-main)", whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis", flex: 1 },
  quickDocTools: { display: "flex", alignItems: "center", gap: 6, flexShrink: 0 },
  pillOk: { fontSize: 9, background: "var(--skin-bg-muted)", padding: "1px 4px", borderRadius: 10, color: "var(--skin-text-muted)" },
  pillWait: { fontSize: 9, background: "var(--skin-bg-warn)", padding: "1px 4px", borderRadius: 10, color: "var(--skin-text-warn)" },
  toolBtn: { background: "none", border: "none", padding: "2px 4px", color: "var(--skin-text-muted)", fontSize: 10, cursor: "pointer", textDecoration: "underline" },
  chevronBtn: { background: "none", border: "none", padding: 0, color: "var(--skin-text-muted)", cursor: "pointer", fontSize: 12, display: "flex", alignItems: "center", justifyContent: "center", width: 16, height: 16 },
  quickDocSnippet: { fontSize: 10, color: "var(--skin-text-muted)", marginTop: 4, lineHeight: "1.3" },
};

export default QuickDocumentsCard;
