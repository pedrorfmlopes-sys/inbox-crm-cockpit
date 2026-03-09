import React, { useMemo } from "react";
import type { LinkEntry } from "@/api";
import type { OutlookMessageContext } from "@/office";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";

function getEntityLabel(model: string): string {
    if (model === "res.partner") return "Contacto";
    if (model === "crm.lead") return "Lead";
    if (model === "project.task") return "Tarefa";
    if (model === "project.project") return "Projeto";
    if (model === "helpdesk.ticket") return "Ticket";
    return model;
}

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
    const previewRecords = linkedRecords.slice(0, 3);

    return (
        <section style={styles.card}>
            <div style={styles.header}>
                <div>
                    <div style={styles.kicker}>Emails relacionados</div>
                    <div style={styles.title}>Contexto de processo</div>
                </div>
                <button style={styles.ctaBtn} onClick={onOpenExplorer}>
                    <Icons.ArrowRight size={12} />
                    Ver tudo
                </button>
            </div>

            {!currentCtx.conversationId && !currentCtx.internetMessageId ? (
                <PanelState
                    tone="info"
                    title="Sem email selecionado"
                    description="Abre o explorador para navegar manualmente por contacto, lead, tarefa, projeto ou ticket."
                    compact
                />
            ) : !linkedRecords.length ? (
                <PanelState
                    tone="empty"
                    title="Sem contexto ligado"
                    description="Este email ainda nao tem registos ligados. O explorador completo continua disponivel para pesquisa manual."
                    compact
                />
            ) : (
                <div style={styles.content}>
                    <div style={styles.metricRow}>
                        <div style={styles.metricCard}>
                            <div style={styles.metricValue}>{linkedRecords.length}</div>
                            <div style={styles.metricLabel}>registos ligados</div>
                        </div>
                        <div style={styles.subjectBox}>
                            <div style={styles.subjectLabel}>Email atual</div>
                            <div style={styles.subjectValue}>{currentCtx.subject || "Sem assunto"}</div>
                        </div>
                    </div>

                    <div style={styles.previewTitle}>Em destaque</div>
                    <div style={styles.recordList}>
                        {previewRecords.map((link) => (
                            <div key={`${link.model}:${link.recordId ?? link.resId}`} style={styles.recordItem}>
                                <span style={styles.recordType}>{getEntityLabel(link.model)}</span>
                                <span style={styles.recordName}>
                                    {link.recordName || link.name || `#${link.recordId ?? link.resId}`}
                                </span>
                            </div>
                        ))}
                    </div>

                    {linkedRecords.length > previewRecords.length ? (
                        <div style={styles.moreHint}>
                            +{linkedRecords.length - previewRecords.length} registo{linkedRecords.length - previewRecords.length === 1 ? "" : "s"} no explorador completo
                        </div>
                    ) : null}
                </div>
            )}
        </section>
    );
}

const styles: Record<string, React.CSSProperties> = {
    card: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        background: "#FFFFFF",
        padding: "12px",
        display: "grid",
        gap: "12px",
    },
    header: {
        display: "flex",
        justifyContent: "space-between",
        gap: "12px",
        alignItems: "center",
    },
    kicker: {
        fontSize: "10px",
        fontWeight: 700,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    title: {
        fontSize: "14px",
        fontWeight: 700,
        color: "#172B4D",
        marginTop: "4px",
    },
    ctaBtn: {
        border: "1px solid #0052CC",
        background: "#DEEBFF",
        color: "#0747A6",
        borderRadius: "16px",
        padding: "6px 12px",
        fontSize: "11px",
        fontWeight: 700,
        cursor: "pointer",
        display: "inline-flex",
        alignItems: "center",
        gap: "4px",
        flexShrink: 0,
    },
    content: {
        display: "grid",
        gap: "10px",
    },
    metricRow: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(min(100%, 160px), 1fr))",
        gap: "10px",
        alignItems: "stretch",
    },
    metricCard: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        background: "#FAFBFC",
        padding: "10px",
        display: "grid",
        gap: "2px",
        alignContent: "center",
    },
    metricValue: {
        fontSize: "20px",
        fontWeight: 800,
        color: "#172B4D",
    },
    metricLabel: {
        fontSize: "11px",
        color: "#6B778C",
        textTransform: "uppercase",
    },
    subjectBox: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        padding: "10px",
        background: "#FAFBFC",
        display: "grid",
        gap: "4px",
    },
    subjectLabel: {
        fontSize: "10px",
        fontWeight: 700,
        color: "#6B778C",
        textTransform: "uppercase",
    },
    subjectValue: {
        fontSize: "12px",
        color: "#172B4D",
        fontWeight: 600,
        lineHeight: 1.4,
        wordBreak: "break-word",
    },
    previewTitle: {
        fontSize: "11px",
        fontWeight: 700,
        color: "#42526E",
        textTransform: "uppercase",
    },
    recordList: {
        display: "grid",
        gap: "8px",
    },
    recordItem: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        padding: "8px 10px",
        display: "flex",
        gap: "8px",
        alignItems: "flex-start",
        background: "#FAFBFC",
        flexWrap: "wrap",
    },
    recordType: {
        fontSize: "10px",
        fontWeight: 700,
        color: "#0747A6",
        textTransform: "uppercase",
        flexShrink: 0,
    },
    recordName: {
        fontSize: "12px",
        color: "#172B4D",
        fontWeight: 600,
        minWidth: 0,
        whiteSpace: "normal",
        wordBreak: "break-word",
    },
    moreHint: {
        fontSize: "11px",
        color: "#6B778C",
    },
};
