import React, { useEffect, useMemo, useState } from "react";
import { getRelatedEmailContext, type LinkEntry, type LinkGroupEntry, type RelatedEmailEntry } from "@/api";
import type { OutlookMessageContext } from "@/office";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

function dedupeRecords(entries: LinkEntry[]): LinkEntry[] {
  const seen = new Set<string>();
  return (entries || []).filter((entry) => {
    const recordId = Number(entry.recordId || entry.resId || 0);
    const key = `${entry.model || ""}:${recordId}`;
    if (!entry.model || !recordId || seen.has(key)) return false;
    seen.add(key);
    return true;
  });
}

function hasContext(ctx: OutlookMessageContext): boolean {
  return Boolean(ctx.itemId || ctx.internetMessageId || ctx.conversationId);
}

export function RelatedEmailsSummaryCard({
  currentCtx,
  currentLinks,
  onOpenExplorer,
}: {
  currentCtx: OutlookMessageContext;
  currentLinks: LinkEntry[];
  onOpenExplorer: () => void;
}) {
  const linkedRecords = useMemo(() => dedupeRecords(currentLinks), [currentLinks]);
  const [relatedCount, setRelatedCount] = useState(0);
  const [groups, setGroups] = useState<LinkGroupEntry[]>([]);
  const [previewEmails, setPreviewEmails] = useState<RelatedEmailEntry[]>([]);

  useEffect(() => {
    if (!hasContext(currentCtx)) {
      setRelatedCount(0);
      setGroups([]);
      setPreviewEmails([]);
      return;
    }
    getRelatedEmailContext({
      itemId: currentCtx.itemId,
      internetMessageId: currentCtx.internetMessageId,
      conversationId: currentCtx.conversationId,
      subject: currentCtx.subject,
      fromEmail: currentCtx.fromEmail,
      receivedAtIso: currentCtx.receivedDateTimeIso,
      messageDateIso: currentCtx.receivedDateTimeIso,
    }).then((response) => {
      setRelatedCount(Array.isArray(response.emails) ? response.emails.length : 0);
      setGroups(Array.isArray(response.groups) ? response.groups : []);
      setPreviewEmails(Array.isArray(response.emails) ? response.emails.slice(0, 2) : []);
    }).catch(() => {
      setRelatedCount(0);
      setGroups([]);
      setPreviewEmails([]);
    });
  }, [currentCtx.itemId, currentCtx.internetMessageId, currentCtx.conversationId, currentCtx.subject, currentCtx.fromEmail, currentCtx.receivedDateTimeIso]);

  return (
    <section style={styles.card}>
      <div style={styles.header}>
        <div style={styles.headerLead}>
          <div style={styles.kicker}>Emails relacionados</div>
          <div style={styles.title}>Resumo de contexto</div>
        </div>
        <button style={styles.ctaBtn} onClick={onOpenExplorer} title="Abrir explorador completo">
          <Icons.ArrowRight size={12} />
        </button>
      </div>

      {!hasContext(currentCtx) ? (
        <PanelState tone="info" title="Sem email selecionado" description="Abre o explorador para navegar manualmente por contexto, Odoo ou grupos." compact />
      ) : (
        <div style={styles.content}>
          <div style={styles.metricRow}>
            <div style={styles.metricChip}><span style={styles.metricValue}>{relatedCount}</span><span style={styles.metricLabel}>emails</span></div>
            <div style={styles.metricChip}><span style={styles.metricValue}>{linkedRecords.length}</span><span style={styles.metricLabel}>registos</span></div>
            <div style={styles.metricChip}><span style={styles.metricValue}>{groups.filter((group) => group.kind === "custom").length}</span><span style={styles.metricLabel}>grupos</span></div>
          </div>

          <div style={styles.subjectValue}>{currentCtx.subject || "Sem assunto"}</div>

          {previewEmails.length ? (
            <div style={styles.previewList}>
              {previewEmails.map((email) => (
                <div key={`${email.id || email.itemId || email.internetMessageId || email.subject}`} style={styles.previewItem}>
                  <Icons.MessageSquare size={11} />
                  <span style={styles.previewText}>{email.subject || "(sem assunto)"}</span>
                </div>
              ))}
            </div>
          ) : linkedRecords.length ? (
            <div style={styles.previewList}>
              {linkedRecords.slice(0, 2).map((record) => (
                <div key={`${record.model}:${record.recordId || record.resId}`} style={styles.previewItem}>
                  <Icons.Link size={11} />
                  <span style={styles.previewText}>{record.recordName || record.name || `#${record.recordId || record.resId}`}</span>
                </div>
              ))}
            </div>
          ) : (
            <div style={styles.emptyHint}>Sem contexto adicional visivel ainda. O explorador completo continua disponivel.</div>
          )}
        </div>
      )}
    </section>
  );
}

const styles: Record<string, React.CSSProperties> = {
  card: { border: "1px solid #DFE1E6", borderRadius: "6px", background: "#FFFFFF", padding: "10px", display: "grid", gap: "10px" },
  header: { display: "flex", justifyContent: "space-between", gap: "8px", alignItems: "center" },
  headerLead: { display: "grid", gap: "2px", minWidth: 0 },
  kicker: { fontSize: "10px", fontWeight: 700, color: "#6B778C", textTransform: "uppercase", letterSpacing: "0.05em" },
  title: { fontSize: "13px", fontWeight: 700, color: "#172B4D" },
  ctaBtn: { border: "1px solid #0052CC", background: "#DEEBFF", color: "#0747A6", borderRadius: "6px", width: "30px", height: "30px", cursor: "pointer", display: "inline-flex", alignItems: "center", justifyContent: "center", flexShrink: 0 },
  content: { display: "grid", gap: "8px" },
  metricRow: { display: "flex", gap: "6px", flexWrap: "wrap" },
  metricChip: { border: "1px solid #DFE1E6", borderRadius: "999px", padding: "4px 8px", background: "#FAFBFC", display: "inline-flex", alignItems: "center", gap: "5px" },
  metricValue: { fontSize: "12px", fontWeight: 800, color: "#172B4D" },
  metricLabel: { fontSize: "10px", color: "#6B778C", textTransform: "uppercase" },
  subjectValue: { fontSize: "12px", color: "#172B4D", fontWeight: 600, lineHeight: 1.4, wordBreak: "break-word" },
  previewList: { display: "grid", gap: "6px", maxHeight: "92px", overflowY: "auto" },
  previewItem: { display: "flex", alignItems: "center", gap: "6px", fontSize: "11px", color: "#42526E", minWidth: 0 },
  previewText: { minWidth: 0, whiteSpace: "nowrap", overflow: "hidden", textOverflow: "ellipsis" },
  emptyHint: { fontSize: "11px", color: "#6B778C", lineHeight: 1.5 },
};
