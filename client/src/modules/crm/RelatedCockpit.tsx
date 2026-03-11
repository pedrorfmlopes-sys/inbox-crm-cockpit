import React, { useMemo } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import type { LinkEntry } from "@/api";
import { PanelState } from "@/ui/PanelState";
import { openCockpitDialog } from "@/office";
import { RelatedEmailsPanel } from "./RelatedEmailsPanel";

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

export const RelatedCockpit: React.FC = () => {
    const { ctx, bodyText, attachments, meta, links, setMsg, settings, refreshLinks } = useCockpit();

    const linkedRecords = useMemo(() => dedupeRecords(links), [links]);

    async function openEditDialog(model: string, recordId: number) {
        try {
            localStorage.setItem("ic_bridge_body", bodyText || "");
            localStorage.setItem("ic_bridge_atts", JSON.stringify(attachments || []));
        } catch {
            // best-effort transition payload only
        }

        try {
            await openCockpitDialog({
                mode: "edit",
                model,
                recordId: String(recordId),
                conversationId: ctx.conversationId || "",
                internetMessageId: ctx.internetMessageId || "",
                itemId: ctx.itemId || "",
                subject: ctx.subject || "",
                fromEmail: ctx.fromEmail || "",
                fromName: ctx.fromName || "",
                receivedAtIso: ctx.receivedDateTimeIso || "",
                toR: ctx.toRecipients || [],
                ccR: ctx.ccRecipients || [],
            } as any);
            await refreshLinks();
        } catch (error: any) {
            setMsg(error?.message ?? String(error));
        }
    }

    return (
        <div style={styles.container}>
            <section style={styles.hero}>
                <div style={styles.heroCopy}>
                    <div style={styles.kicker}>Contexto</div>
                    <h2 style={styles.title}>Emails relacionados</h2>
                    <p style={styles.description}>
                        Explora contexto Outlook, grupos manuais e processos Odoo num painel unico e compacto.
                    </p>
                </div>

                <div style={styles.heroStats}>
                    <div style={styles.metricChip}>
                        <span style={styles.metricValue}>{linkedRecords.length}</span>
                        <span style={styles.metricLabel}>registos</span>
                    </div>
                    <div style={styles.subjectChip}>
                        <div style={styles.heroLabel}>Email atual</div>
                        <div style={styles.heroSubject}>{ctx.subject || "Nenhum email selecionado"}</div>
                    </div>
                </div>
            </section>

            {!ctx.conversationId && !ctx.internetMessageId ? (
                <PanelState
                    tone="info"
                    title="Sem contexto do email atual"
                    description="O explorador manual continua disponivel neste tab, mesmo quando nao existe uma conversa aberta."
                />
            ) : null}

            {linkedRecords.length ? (
                <section style={styles.contextSummary}>
                    <div style={styles.summaryTitle}>Registos ligados ao email aberto</div>
                    <div style={styles.recordGrid}>
                        {linkedRecords.map((link) => (
                            <div key={`${link.model}:${link.recordId ?? link.resId}`} style={styles.recordChip}>
                                <span style={styles.recordType}>{getEntityLabel(link.model)}</span>
                                <span style={styles.recordName}>
                                    {link.recordName || link.name || `#${link.recordId ?? link.resId}`}
                                </span>
                            </div>
                        ))}
                    </div>
                </section>
            ) : null}

            <RelatedEmailsPanel
                currentCtx={ctx}
                currentLinks={links}
                meta={meta}
                settings={settings}
                onEditRecord={openEditDialog}
                onStatus={(message) => setMsg(message)}
            />
        </div>
    );
};

const styles: Record<string, React.CSSProperties> = {
    container: {
        display: "grid",
        gap: "8px",
        alignContent: "start",
    },
    hero: {
        display: "grid",
        gap: "8px",
        border: "1px solid #DFE1E6",
        borderRadius: "10px",
        background: "#FFFFFF",
        padding: "8px 10px",
    },
    heroCopy: {
        display: "grid",
        gap: "4px",
        minWidth: 0,
    },
    kicker: {
        fontSize: "10px",
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    title: {
        margin: 0,
        fontSize: "15px",
        lineHeight: 1.2,
        color: "#172B4D",
    },
    description: {
        margin: 0,
        fontSize: "11px",
        lineHeight: 1.35,
        color: "#42526E",
    },
    heroStats: {
        display: "grid",
        gridTemplateColumns: "88px minmax(0, 1fr)",
        gap: "6px",
        minWidth: 0,
    },
    metricChip: {
        border: "1px solid #DFE1E6",
        borderRadius: "999px",
        background: "#FAFBFC",
        padding: "4px 8px",
        display: "grid",
        alignItems: "center",
        justifyItems: "start",
        gap: "1px",
        minWidth: 0,
    },
    metricValue: {
        fontSize: "13px",
        fontWeight: 800,
        color: "#172B4D",
    },
    metricLabel: {
        fontSize: "9px",
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.04em",
    },
    subjectChip: {
        border: "1px solid #DFE1E6",
        borderRadius: "10px",
        background: "#FAFBFC",
        padding: "6px 8px",
        display: "grid",
        gap: "2px",
        minWidth: 0,
    },
    heroLabel: {
        fontSize: "9px",
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    heroSubject: {
        fontSize: "11px",
        fontWeight: 600,
        color: "#172B4D",
        lineHeight: 1.3,
        wordBreak: "break-word",
    },
    contextSummary: {
        border: "1px solid #DFE1E6",
        borderRadius: "10px",
        background: "#FFFFFF",
        padding: "8px 10px",
        display: "grid",
        gap: "6px",
    },
    summaryTitle: {
        fontSize: "10px",
        fontWeight: 800,
        color: "#42526E",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    recordGrid: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
        maxHeight: "86px",
        overflowY: "auto",
        paddingRight: "2px",
    },
    recordChip: {
        border: "1px solid #DFE1E6",
        borderRadius: "999px",
        padding: "4px 8px",
        background: "#FAFBFC",
        display: "inline-flex",
        alignItems: "center",
        gap: "5px",
        minWidth: 0,
        maxWidth: "100%",
    },
    recordType: {
        fontSize: "9px",
        fontWeight: 800,
        color: "#0747A6",
        textTransform: "uppercase",
        flexShrink: 0,
    },
    recordName: {
        fontSize: "11px",
        color: "#172B4D",
        whiteSpace: "normal",
        wordBreak: "break-word",
    },
};
