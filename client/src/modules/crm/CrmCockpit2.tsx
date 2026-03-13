import React, { useEffect, useMemo, useState } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import { openCockpitDialog } from "@/office";
import { getOdooAutoLoginUrl, getPartnerByEmail, type LinkEntry } from "@/api";

type Participant = {
    email: string;
    name: string;
    source: "from" | "to" | "cc";
};

function normalizeEmail(value: string | undefined): string {
    return String(value || "").trim().toLowerCase();
}

function fallbackNameFromEmail(email: string): string {
    const local = String(email || "").split("@")[0] || "Contacto";
    return local
        .replace(/[._-]+/g, " ")
        .replace(/\b\w/g, (match) => match.toUpperCase())
        .trim();
}

function getModelLabel(model: string): string {
    if (model === "res.partner") return "Contacto";
    if (model === "crm.lead") return "Lead";
    if (model === "project.project") return "Projeto";
    if (model === "project.task") return "Tarefa";
    if (model === "helpdesk.ticket") return "Ticket";
    return model;
}

function dedupeParticipants(ctx: any): Participant[] {
    const seen = new Set<string>();
    const rows: Participant[] = [];
    const push = (source: Participant["source"], email?: string, name?: string) => {
        const key = normalizeEmail(email);
        if (!key || seen.has(key)) return;
        seen.add(key);
        rows.push({
            email: key,
            name: String(name || "").trim() || fallbackNameFromEmail(key),
            source,
        });
    };

    push("from", ctx.fromEmail, ctx.fromName);
    for (const recipient of ctx.toRecipients || []) push("to", recipient?.email, recipient?.name);
    for (const recipient of ctx.ccRecipients || []) push("cc", recipient?.email, recipient?.name);
    return rows;
}

function dedupeLinkedRecords(entries: LinkEntry[]): LinkEntry[] {
    const seen = new Set<string>();
    return (entries || []).filter((entry) => {
        const id = Number(entry.recordId || entry.resId || 0);
        const key = `${entry.model}:${id}`;
        if (!entry.model || !id || seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function serializeRecipients(recipients: Array<{ name?: string; email?: string }> | undefined) {
    return (recipients || [])
        .map((recipient) => `${String(recipient?.name || "").trim()}|${String(recipient?.email || "").trim()}`)
        .filter((value) => !value.endsWith("|"))
        .join(";");
}

const QUICK_CREATE_MODELS = [
    { model: "project.project", label: "Projeto", icon: <Icons.Database size={12} /> },
    { model: "crm.lead", label: "Lead", icon: <Icons.Target size={12} /> },
    { model: "project.task", label: "Tarefa", icon: <Icons.Clipboard size={12} /> },
    { model: "helpdesk.ticket", label: "Ticket", icon: <Icons.MessageSquare size={12} /> },
    { model: "res.partner", label: "Contacto", icon: <Icons.User size={12} /> },
] as const;

export const CrmCockpit2: React.FC = () => {
    const { ctx, bodyText, bodyHtml, attachments, links, settings, meta, refreshLinks, setMsg, setTab } = useCockpit();
    const [primaryPartner, setPrimaryPartner] = useState<any | null>(null);
    const [contactLoading, setContactLoading] = useState(false);
    const [participantsExpanded, setParticipantsExpanded] = useState(false);
    const [linkedExpanded, setLinkedExpanded] = useState(true);

    const participants = useMemo(() => dedupeParticipants(ctx), [ctx]);
    const linkedRecords = useMemo(() => dedupeLinkedRecords(links), [links]);
    const linkedGroups = useMemo(() => {
        const names = new Set<string>();
        for (const link of links || []) {
            const name = String((link as any).groupName || "").trim();
            if (name) names.add(name);
        }
        return Array.from(names).sort((a, b) => a.localeCompare(b));
    }, [links]);

    useEffect(() => {
        let alive = true;
        const email = normalizeEmail(ctx.fromEmail);
        if (!email) {
            setPrimaryPartner(null);
            return;
        }
        setContactLoading(true);
        getPartnerByEmail(email)
            .then((partner) => {
                if (!alive) return;
                setPrimaryPartner(partner || null);
            })
            .catch(() => {
                if (!alive) return;
                setPrimaryPartner(null);
            })
            .finally(() => {
                if (alive) setContactLoading(false);
            });
        return () => {
            alive = false;
        };
    }, [ctx.fromEmail]);

    async function openDialog(targetMode: "new" | "add" | "edit", extra?: Record<string, string>) {
        if (!ctx.conversationId && targetMode !== "edit") {
            setMsg("Seleciona um email primeiro.");
            return;
        }

        try {
            localStorage.setItem("ic_bridge_body", bodyText || "");
            localStorage.setItem("ic_bridge_html", bodyHtml || "");
            localStorage.setItem("ic_bridge_atts", JSON.stringify(attachments || []));
        } catch {
            // best effort only
        }

        const useV2 = targetMode !== "add";
        try {
            await openCockpitDialog({
                mode: targetMode,
                conversationId: ctx.conversationId || "",
                internetMessageId: ctx.internetMessageId || "",
                itemId: ctx.itemId || "",
                subject: ctx.subject || "",
                fromEmail: ctx.fromEmail || "",
                fromName: ctx.fromName || "",
                receivedAtIso: ctx.receivedDateTimeIso || "",
                emailWebLink: (ctx as any).emailWebLink || "",
                toR: serializeRecipients(ctx.toRecipients || []),
                ccR: serializeRecipients(ctx.ccRecipients || []),
                ...(useV2 ? { ui: "v2" } : {}),
                ...(extra ?? {}),
            } as any);
            await refreshLinks();
        } catch (error: any) {
            setMsg(error?.message ?? String(error));
        }
    }

    function openOdooRecord(model: string, recordId?: number | null) {
        const baseUrl = settings?.odooUrl || meta?.baseUrl || meta?.webBaseUrl || meta?.url || "";
        const db = settings?.odooDb || meta?.db || "";
        const targetBase = db ? `/web?db=${encodeURIComponent(db)}` : "/web";
        const target = recordId
            ? `${targetBase}#id=${encodeURIComponent(String(recordId))}&model=${encodeURIComponent(model)}&view_type=form`
            : targetBase;
        window.open(getOdooAutoLoginUrl(settings?.odooSessionToken || null, target, baseUrl), "_blank");
    }

    const primaryName = primaryPartner?.name || ctx.fromName || fallbackNameFromEmail(ctx.fromEmail || "");
    const primaryCompany =
        primaryPartner?.parent_id?.[1] ||
        primaryPartner?.company_name ||
        (primaryPartner?.company_type === "company" ? primaryPartner?.name : "");
    const primaryRole = primaryPartner?.function || "";

    return (
        <div style={S.root}>
            <section style={S.hero}>
                <div style={S.heroHead}>
                    <div>
                        <div style={S.kicker}>CRM 2</div>
                        <h2 style={S.title}>Launcher compacto de CRM</h2>
                        <p style={S.copy}>
                            Foco no contacto principal, registos ligados e acoes rapidas. O detalhe fica no novo editor.
                        </p>
                    </div>
                    <button type="button" style={S.secondaryAction} onClick={() => setTab("crm")}>
                        Abrir CRM atual
                    </button>
                </div>

                <div style={S.heroGrid}>
                    <div style={S.primaryCard}>
                        <div style={S.cardLabel}>Contacto principal</div>
                        <div style={S.primaryName}>{primaryName || "Sem contacto identificado"}</div>
                        <div style={S.primaryMeta}>{ctx.fromEmail || "Sem email"}</div>
                        {primaryRole ? <div style={S.primaryMinor}>{primaryRole}</div> : null}
                        {primaryCompany ? <div style={S.primaryMinor}>{primaryCompany}</div> : null}
                        <div style={S.contactActions}>
                            <button
                                type="button"
                                style={S.primaryAction}
                                onClick={() => {
                                    if (primaryPartner?.id) openDialog("edit", { model: "res.partner", recordId: String(primaryPartner.id) });
                                    else openDialog("new", { model: "res.partner" });
                                }}
                            >
                                {primaryPartner?.id ? "Editar contacto" : "Criar contacto"}
                            </button>
                            <button
                                type="button"
                                style={S.secondaryAction}
                                onClick={() => openOdooRecord("res.partner", Number(primaryPartner?.id || 0) || undefined)}
                                disabled={!settings?.odooUrl && !meta?.baseUrl && !meta?.webBaseUrl && !meta?.url}
                            >
                                Odoo
                            </button>
                        </div>
                    </div>

                    <div style={S.subjectCard}>
                        <div style={S.cardLabel}>Email atual</div>
                        <div style={S.subjectText} title={ctx.subject || ""}>{ctx.subject || "Sem assunto"}</div>
                        <div style={S.subjectMeta}>
                            <span>{participants.length} participante(s)</span>
                            <span>{linkedRecords.length} registo(s)</span>
                            <span>{linkedGroups.length} grupo(s)</span>
                        </div>
                    </div>
                </div>
            </section>

            <section style={S.quickCreateCard}>
                <div style={S.sectionHead}>
                    <div>
                        <div style={S.sectionTitle}>Criar ou ligar</div>
                        <div style={S.sectionHint}>Abre o novo editor para criar ou o fluxo atual para ligar um registo existente.</div>
                    </div>
                    <button type="button" style={S.secondaryAction} onClick={() => openDialog("add")}>
                        Ligar existente
                    </button>
                </div>

                <div style={S.quickCreateGrid}>
                    {QUICK_CREATE_MODELS.map((item) => (
                        <button
                            key={item.model}
                            type="button"
                            style={S.quickCreateBtn}
                            onClick={() => openDialog("new", { model: item.model })}
                        >
                            <span style={S.quickCreateIcon}>{item.icon}</span>
                            <span>{item.label}</span>
                        </button>
                    ))}
                </div>
            </section>

            <section style={S.sectionCard}>
                <div style={S.sectionHead}>
                    <div>
                        <div style={S.sectionTitle}>Participantes</div>
                        <div style={S.sectionHint}>Mostra o contacto principal e os restantes participantes do email.</div>
                    </div>
                    <button type="button" style={S.secondaryAction} onClick={() => setParticipantsExpanded((value) => !value)}>
                        {participantsExpanded ? "Recolher" : "Expandir"}
                    </button>
                </div>

                {contactLoading ? (
                    <PanelState compact tone="loading" title="A procurar contacto" description="A cruzar o remetente com o Odoo." />
                ) : participantsExpanded ? (
                    <div style={S.participantList}>
                        {participants.map((row) => {
                            const isPrimary = normalizeEmail(row.email) === normalizeEmail(ctx.fromEmail);
                            return (
                                <div key={`${row.source}:${row.email}`} style={isPrimary ? { ...S.participantCard, ...S.participantCardPrimary } : S.participantCard}>
                                    <div>
                                        <div style={S.participantName}>{row.name}</div>
                                        <div style={S.participantEmail}>{row.email}</div>
                                    </div>
                                    <span style={S.participantBadge}>{row.source.toUpperCase()}</span>
                                </div>
                            );
                        })}
                    </div>
                ) : (
                    <div style={S.collapsedSummary}>
                        {participants.map((row) => row.name).slice(0, 4).join(" · ")}
                        {participants.length > 4 ? ` +${participants.length - 4}` : ""}
                    </div>
                )}
            </section>

            <section style={S.sectionCard}>
                <div style={S.sectionHead}>
                    <div>
                        <div style={S.sectionTitle}>Registos ligados</div>
                        <div style={S.sectionHint}>Leitura direta do storage central dos links desta conversa.</div>
                    </div>
                    <div style={S.sectionHeadActions}>
                        <button type="button" style={S.secondaryAction} onClick={() => refreshLinks()}>
                            Atualizar
                        </button>
                        <button type="button" style={S.secondaryAction} onClick={() => setLinkedExpanded((value) => !value)}>
                            {linkedExpanded ? "Recolher" : "Expandir"}
                        </button>
                    </div>
                </div>

                {!linkedRecords.length ? (
                    <PanelState
                        compact
                        tone="empty"
                        title="Sem registos ligados"
                        description="Usa criar ou ligar para associar esta conversa a um contacto, lead, projeto, tarefa ou ticket."
                    />
                ) : linkedExpanded ? (
                    <div style={S.linkedGrid}>
                        {linkedRecords.map((link) => {
                            const id = Number(link.recordId || link.resId || 0);
                            return (
                                <div key={`${link.model}:${id}`} style={S.linkedCard}>
                                    <div style={S.linkedType}>{getModelLabel(link.model)}</div>
                                    <div style={S.linkedName}>{link.recordName || link.name || `#${id}`}</div>
                                    <div style={S.linkedActions}>
                                        <button type="button" style={S.linkAction} onClick={() => openDialog("edit", { model: link.model, recordId: String(id) })}>
                                            Editor v2
                                        </button>
                                        <button type="button" style={S.linkActionMuted} onClick={() => openOdooRecord(link.model, id)}>
                                            Odoo
                                        </button>
                                    </div>
                                </div>
                            );
                        })}
                    </div>
                ) : (
                    <div style={S.collapsedSummary}>
                        {linkedRecords.map((link) => `${getModelLabel(link.model)}: ${link.recordName || link.name || "#"}`).slice(0, 3).join(" · ")}
                    </div>
                )}
            </section>

            <section style={S.sectionCard}>
                <div style={S.sectionHead}>
                    <div>
                        <div style={S.sectionTitle}>Contexto rapido</div>
                        <div style={S.sectionHint}>Usa o tab Contexto para navegar emails relacionados e grupos manuais.</div>
                    </div>
                    <button type="button" style={S.primaryAction} onClick={() => setTab("related")}>
                        Abrir Contexto
                    </button>
                </div>
            </section>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    root: {
        display: "grid",
        gap: 8,
        alignContent: "start",
    },
    hero: {
        display: "grid",
        gap: 8,
        border: "1px solid #DFE1E6",
        borderRadius: 12,
        background: "#FFFFFF",
        padding: "10px",
    },
    heroHead: {
        display: "flex",
        alignItems: "start",
        justifyContent: "space-between",
        gap: 10,
    },
    kicker: {
        fontSize: 10,
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    title: {
        margin: "2px 0 0",
        fontSize: 16,
        lineHeight: 1.2,
        color: "#172B4D",
    },
    copy: {
        margin: "4px 0 0",
        fontSize: 11,
        lineHeight: 1.35,
        color: "#42526E",
        maxWidth: 480,
    },
    heroGrid: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
        gap: 8,
    },
    primaryCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "grid",
        gap: 4,
    },
    subjectCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "grid",
        gap: 6,
    },
    cardLabel: {
        fontSize: 10,
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    primaryName: {
        fontSize: 16,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.2,
    },
    primaryMeta: {
        fontSize: 12,
        fontWeight: 600,
        color: "#253858",
    },
    primaryMinor: {
        fontSize: 11,
        color: "#5E6C84",
    },
    subjectText: {
        fontSize: 13,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.35,
        display: "-webkit-box",
        WebkitLineClamp: 2,
        WebkitBoxOrient: "vertical",
        overflow: "hidden",
    },
    subjectMeta: {
        display: "flex",
        flexWrap: "wrap",
        gap: 6,
        fontSize: 10,
        color: "#5E6C84",
    },
    contactActions: {
        display: "flex",
        gap: 8,
        marginTop: 4,
        flexWrap: "wrap",
    },
    primaryAction: {
        border: "none",
        borderRadius: 8,
        background: "#0C66E4",
        color: "#FFFFFF",
        padding: "7px 12px",
        fontSize: 11,
        fontWeight: 700,
        cursor: "pointer",
    },
    secondaryAction: {
        border: "1px solid #DFE1E6",
        borderRadius: 8,
        background: "#FFFFFF",
        color: "#42526E",
        padding: "7px 12px",
        fontSize: 11,
        fontWeight: 700,
        cursor: "pointer",
    },
    quickCreateCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 12,
        background: "#FFFFFF",
        padding: "10px",
        display: "grid",
        gap: 10,
    },
    sectionCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 12,
        background: "#FFFFFF",
        padding: "10px",
        display: "grid",
        gap: 10,
    },
    sectionHead: {
        display: "flex",
        alignItems: "start",
        justifyContent: "space-between",
        gap: 10,
    },
    sectionHeadActions: {
        display: "flex",
        gap: 8,
        flexWrap: "wrap",
    },
    sectionTitle: {
        fontSize: 13,
        fontWeight: 800,
        color: "#172B4D",
    },
    sectionHint: {
        marginTop: 2,
        fontSize: 10,
        lineHeight: 1.35,
        color: "#5E6C84",
        maxWidth: 420,
    },
    quickCreateGrid: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(92px, 1fr))",
        gap: 8,
    },
    quickCreateBtn: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 8px",
        display: "grid",
        justifyItems: "center",
        gap: 6,
        fontSize: 11,
        fontWeight: 700,
        color: "#172B4D",
        cursor: "pointer",
    },
    quickCreateIcon: {
        color: "#0C66E4",
        display: "inline-flex",
    },
    participantList: {
        display: "grid",
        gap: 8,
    },
    participantCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        gap: 12,
    },
    participantCardPrimary: {
        borderColor: "#0C66E4",
        background: "#EFF6FF",
    },
    participantName: {
        fontSize: 12,
        fontWeight: 700,
        color: "#172B4D",
    },
    participantEmail: {
        fontSize: 11,
        color: "#42526E",
    },
    participantBadge: {
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        minWidth: 36,
        borderRadius: 999,
        padding: "3px 8px",
        fontSize: 10,
        fontWeight: 800,
        background: "#DFE1E6",
        color: "#42526E",
    },
    collapsedSummary: {
        fontSize: 11,
        color: "#42526E",
        lineHeight: 1.45,
    },
    linkedGrid: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(180px, 1fr))",
        gap: 8,
    },
    linkedCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "grid",
        gap: 8,
    },
    linkedType: {
        fontSize: 10,
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    linkedName: {
        fontSize: 12,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.35,
    },
    linkedActions: {
        display: "flex",
        gap: 8,
        flexWrap: "wrap",
    },
    linkAction: {
        border: "none",
        borderRadius: 8,
        background: "#0C66E4",
        color: "#FFFFFF",
        padding: "6px 10px",
        fontSize: 10,
        fontWeight: 700,
        cursor: "pointer",
    },
    linkActionMuted: {
        border: "1px solid #DFE1E6",
        borderRadius: 8,
        background: "#FFFFFF",
        color: "#42526E",
        padding: "6px 10px",
        fontSize: 10,
        fontWeight: 700,
        cursor: "pointer",
    },
};
