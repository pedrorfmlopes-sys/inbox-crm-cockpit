import React, { useEffect, useMemo, useState } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { PanelState } from "@/ui/PanelState";
import * as Icons from "@/ui/icons";
import { openCockpitDialog } from "@/office";
import {
    getLinksByRecord,
    getOdooAutoLoginUrl,
    getPartnerByEmail,
    getPartnerRelations,
    linkEmailToRecord,
    type LinkEntry,
    type PartnerRelationItem,
    type PartnerRelationSection,
} from "@/api";

type Participant = {
    email: string;
    name: string;
    source: "from" | "to" | "cc";
};

type ParticipantCollectionState =
    | { kind: "relation"; key: string }
    | { kind: "storage" }
    | null;

type ParticipantCollectionSort = "recent" | "title_asc" | "title_desc";
type ParticipantCollectionFilter = "all" | "title" | "meta" | "detail";

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

function dedupeParticipantLinks(entries: LinkEntry[]): LinkEntry[] {
    const seen = new Set<string>();
    return (entries || []).filter((entry) => {
        const key = [
            String(entry.conversationId || "").trim(),
            String(entry.itemId || "").trim(),
            String(entry.emailWebLink || entry.url || "").trim(),
            String(entry.subject || "").trim(),
        ].join("|");
        if (seen.has(key)) return false;
        seen.add(key);
        return true;
    });
}

function sortLinksByRecency(entries: LinkEntry[]): LinkEntry[] {
    return [...entries].sort((a, b) => {
        const aTime = Date.parse(String(a.linkedAt || a.receivedAtIso || a.messageDateIso || a.sentAtIso || 0));
        const bTime = Date.parse(String(b.linkedAt || b.receivedAtIso || b.messageDateIso || b.sentAtIso || 0));
        return (Number.isFinite(bTime) ? bTime : 0) - (Number.isFinite(aTime) ? aTime : 0);
    });
}

function formatLinkMoment(link: LinkEntry): string {
    const raw = String(link.linkedAt || link.receivedAtIso || link.messageDateIso || link.sentAtIso || "").trim();
    if (!raw) return "Sem data";
    const date = new Date(raw);
    if (Number.isNaN(date.getTime())) return "Sem data";
    return date.toLocaleString("pt-PT", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
    });
}

function normalizeSearchText(value: string | undefined): string {
    return String(value || "").trim().toLowerCase();
}

function relationItemFieldText(item: PartnerRelationItem, field: ParticipantCollectionFilter): string {
    if (field === "title") return String(item.title || "");
    if (field === "meta") return [item.meta, item.secondary].filter(Boolean).join(" ");
    if (field === "detail") return [item.title, item.meta, item.secondary, item.model, item.recordId].filter(Boolean).join(" ");
    return [item.title, item.meta, item.secondary, item.model, item.recordId].filter(Boolean).join(" ");
}

function linkFieldText(entry: LinkEntry, field: ParticipantCollectionFilter): string {
    if (field === "title") return String(entry.subject || "");
    if (field === "meta") return [entry.fromName, entry.fromEmail].filter(Boolean).join(" ");
    if (field === "detail") return [entry.emailWebLink, entry.url, entry.conversationId, entry.itemId].filter(Boolean).join(" ");
    return [entry.subject, entry.fromEmail, entry.fromName, entry.emailWebLink, entry.url, entry.conversationId, entry.itemId].filter(Boolean).join(" ");
}

function summarizeRelationSection(section: PartnerRelationSection): string {
    if (!section.items.length) return "Sem registos nesta colecao.";
    return section.items
        .slice(0, 2)
        .map((item) => item.title)
        .filter(Boolean)
        .join(" · ");
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

const CRM2_EDITABLE_MODELS = new Set(["res.partner", "crm.lead", "project.project", "project.task", "helpdesk.ticket"]);

function canOpenCrm2Editor(model: string): boolean {
    return CRM2_EDITABLE_MODELS.has(String(model || "").trim());
}

export const CrmCockpit2: React.FC = () => {
    const { ctx, bodyText, bodyHtml, attachments, links, settings, meta, refreshLinks, setMsg, setTab } = useCockpit();
    const [primaryPartner, setPrimaryPartner] = useState<any | null>(null);
    const [contactLoading, setContactLoading] = useState(false);
    const [participantsExpanded, setParticipantsExpanded] = useState(false);
    const [linkedExpanded, setLinkedExpanded] = useState(true);
    const [activeParticipantEmail, setActiveParticipantEmail] = useState<string | null>(null);
    const [participantDetailLoading, setParticipantDetailLoading] = useState(false);
    const [participantDetailError, setParticipantDetailError] = useState("");
    const [participantDetailPartner, setParticipantDetailPartner] = useState<any | null>(null);
    const [participantDetailLinks, setParticipantDetailLinks] = useState<LinkEntry[]>([]);
    const [participantDetailRelations, setParticipantDetailRelations] = useState<PartnerRelationSection[]>([]);
    const [activeParticipantCollection, setActiveParticipantCollection] = useState<ParticipantCollectionState>(null);
    const [participantCollectionQuery, setParticipantCollectionQuery] = useState("");
    const [participantCollectionSort, setParticipantCollectionSort] = useState<ParticipantCollectionSort>("recent");
    const [participantCollectionFilter, setParticipantCollectionFilter] = useState<ParticipantCollectionFilter>("all");
    const [participantActionBusy, setParticipantActionBusy] = useState(false);

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
    const activeParticipant = useMemo(
        () => participants.find((row) => normalizeEmail(row.email) === normalizeEmail(activeParticipantEmail || "")) || null,
        [participants, activeParticipantEmail],
    );

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

    async function loadParticipantDetail(participant: Participant) {
        setParticipantDetailLoading(true);
        setParticipantDetailError("");
        try {
            const partner = await getPartnerByEmail(participant.email);
            const [recordLinks, relationPayload] = partner?.id
                ? await Promise.all([
                    getLinksByRecord("res.partner", Number(partner.id)),
                    getPartnerRelations(Number(partner.id)),
                ])
                : [[], { partner: null, total: 0, relations: [] }];
            setParticipantDetailPartner(partner || null);
            setParticipantDetailLinks(sortLinksByRecency(dedupeParticipantLinks(recordLinks)));
            setParticipantDetailRelations(Array.isArray(relationPayload?.relations) ? relationPayload.relations : []);
        } catch (error: any) {
            setParticipantDetailPartner(null);
            setParticipantDetailLinks([]);
            setParticipantDetailRelations([]);
            setParticipantDetailError(error?.message ?? String(error));
        } finally {
            setParticipantDetailLoading(false);
        }
    }

    async function openParticipantDetail(participant: Participant) {
        setActiveParticipantEmail(participant.email);
        setActiveParticipantCollection(null);
        setParticipantCollectionQuery("");
        setParticipantCollectionSort("recent");
        setParticipantCollectionFilter("all");
        await loadParticipantDetail(participant);
    }

    function openParticipantCollection(collection: ParticipantCollectionState) {
        setActiveParticipantCollection(collection);
        setParticipantCollectionQuery("");
        setParticipantCollectionSort(collection?.kind === "storage" ? "recent" : "title_asc");
        setParticipantCollectionFilter("all");
    }

    async function openParticipantEditor(participant: Participant, partnerId?: number | null) {
        await openDialog(partnerId ? "edit" : "new", {
            model: "res.partner",
            fromEmail: participant.email,
            fromName: participant.name,
            ...(partnerId ? { recordId: String(partnerId) } : {}),
        });
        await loadParticipantDetail(participant);
    }

    async function linkParticipantToCurrentEmail(participant: Participant) {
        if (!participantDetailPartner?.id) return;
        setParticipantActionBusy(true);
        try {
            await linkEmailToRecord({
                conversationId: ctx.conversationId,
                model: "res.partner",
                recordId: Number(participantDetailPartner.id),
                recordName: participantDetailPartner.name || participant.name || participant.email,
                internetMessageId: ctx.internetMessageId,
                itemId: ctx.itemId,
                subject: ctx.subject,
                fromEmail: ctx.fromEmail,
                fromName: ctx.fromName,
                receivedAtIso: ctx.receivedDateTimeIso,
                emailWebLink: (ctx as any).emailWebLink,
            });
            await refreshLinks();
            await loadParticipantDetail(participant);
            setMsg("Contacto ligado ao email atual.");
        } catch (error: any) {
            setMsg(error?.message ?? String(error));
        } finally {
            setParticipantActionBusy(false);
        }
    }

    useEffect(() => {
        if (!activeParticipantEmail) return;
        const stillVisible = participants.some((row) => normalizeEmail(row.email) === normalizeEmail(activeParticipantEmail));
        if (stillVisible) return;
        setActiveParticipantEmail(null);
        setActiveParticipantCollection(null);
        setParticipantDetailPartner(null);
        setParticipantDetailLinks([]);
        setParticipantDetailRelations([]);
        setParticipantCollectionQuery("");
        setParticipantCollectionSort("recent");
        setParticipantCollectionFilter("all");
        setParticipantDetailError("");
    }, [participants, activeParticipantEmail]);

    useEffect(() => {
        setActiveParticipantEmail(null);
        setActiveParticipantCollection(null);
        setParticipantDetailPartner(null);
        setParticipantDetailLinks([]);
        setParticipantDetailRelations([]);
        setParticipantCollectionQuery("");
        setParticipantCollectionSort("recent");
        setParticipantCollectionFilter("all");
        setParticipantDetailError("");
    }, [ctx.conversationId, ctx.itemId]);

    const primaryName = primaryPartner?.name || ctx.fromName || fallbackNameFromEmail(ctx.fromEmail || "");
    const primaryCompany =
        primaryPartner?.parent_id?.[1] ||
        primaryPartner?.company_name ||
        (primaryPartner?.company_type === "company" ? primaryPartner?.name : "");
    const primaryRole = primaryPartner?.function || "";
    const participantCompany =
        participantDetailPartner?.parent_id?.[1] ||
        participantDetailPartner?.company_name ||
        (participantDetailPartner?.company_type === "company" ? participantDetailPartner?.name : "");
    const participantCurrentConversationLinked = participantDetailPartner?.id
        ? linkedRecords.some((entry) => entry.model === "res.partner" && Number(entry.recordId || entry.resId || 0) === Number(participantDetailPartner.id))
        : false;
    const participantRelationTotal = participantDetailRelations.reduce((sum, section) => sum + Number(section?.total || 0), 0);
    const participantRelationPreview = participantDetailRelations
        .slice(0, 2)
        .map((section) => `${section.total} ${section.label.toLowerCase()}`)
        .join(" · ");
    const activeRelationSection = useMemo(
        () => activeParticipantCollection?.kind === "relation"
            ? participantDetailRelations.find((section) => section.key === activeParticipantCollection.key) || null
            : null,
        [activeParticipantCollection, participantDetailRelations],
    );
    const collectionQuery = normalizeSearchText(participantCollectionQuery);
    const filteredRelationItems = useMemo(() => {
        if (!activeRelationSection) return [];
        const rows = [...activeRelationSection.items];
        const filtered = collectionQuery
            ? rows.filter((item) => normalizeSearchText(relationItemFieldText(item, participantCollectionFilter)).includes(collectionQuery))
            : rows;
        if (participantCollectionSort === "title_desc") {
            filtered.sort((a, b) => String(b.title || "").localeCompare(String(a.title || ""), "pt-PT"));
        } else {
            filtered.sort((a, b) => String(a.title || "").localeCompare(String(b.title || ""), "pt-PT"));
        }
        return filtered;
    }, [activeRelationSection, collectionQuery, participantCollectionFilter, participantCollectionSort]);
    const filteredStorageLinks = useMemo(() => {
        const rows = [...participantDetailLinks];
        const filtered = collectionQuery
            ? rows.filter((entry) => normalizeSearchText(linkFieldText(entry, participantCollectionFilter)).includes(collectionQuery))
            : rows;
        if (participantCollectionSort === "title_asc") {
            filtered.sort((a, b) => String(a.subject || "").localeCompare(String(b.subject || ""), "pt-PT"));
        } else if (participantCollectionSort === "title_desc") {
            filtered.sort((a, b) => String(b.subject || "").localeCompare(String(a.subject || ""), "pt-PT"));
        } else {
            filtered.sort((a, b) => {
                const aTime = Date.parse(String(a.linkedAt || a.receivedAtIso || a.messageDateIso || a.sentAtIso || 0));
                const bTime = Date.parse(String(b.linkedAt || b.receivedAtIso || b.messageDateIso || b.sentAtIso || 0));
                return (Number.isFinite(bTime) ? bTime : 0) - (Number.isFinite(aTime) ? aTime : 0);
            });
        }
        return filtered;
    }, [participantDetailLinks, collectionQuery, participantCollectionFilter, participantCollectionSort]);
    const participantTrackTransform = activeParticipantCollection
        ? "translateX(-66.6667%)"
        : activeParticipant
            ? "translateX(-33.3333%)"
            : "translateX(0)";
    const participantCollectionTitle = activeParticipantCollection?.kind === "relation"
        ? activeRelationSection?.label || "Colecao"
        : activeParticipantCollection?.kind === "storage"
            ? "Emails ligados"
            : "";
    const participantCollectionSubtitle = activeParticipantCollection?.kind === "relation"
        ? `${activeParticipant?.name || ""} - ${activeRelationSection?.total || 0} registo(s)`
        : activeParticipantCollection?.kind === "storage"
            ? `${activeParticipant?.name || ""} - ${participantDetailLinks.length} email(s)`
            : "";

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

                <div
                    style={{
                        ...S.heroGrid,
                        gridTemplateColumns: activeParticipant ? "minmax(0, 1fr)" : S.heroGrid.gridTemplateColumns,
                    }}
                >
                    {!activeParticipant ? (
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
                    ) : null}

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

            {!activeParticipant ? (
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
            ) : null}

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

                <div style={S.participantViewport}>
                    <div
                        style={{
                            ...S.participantTrack,
                            transform: participantTrackTransform,
                        }}
                    >
                        <div style={S.participantPane}>
                            {contactLoading ? (
                                <PanelState compact tone="loading" title="A procurar contacto" description="A cruzar o remetente com o Odoo." />
                            ) : participantsExpanded ? (
                    <div style={S.participantList}>
                        {participants.map((row) => {
                            const isPrimary = normalizeEmail(row.email) === normalizeEmail(ctx.fromEmail);
                            const isActive = normalizeEmail(row.email) === normalizeEmail(activeParticipant?.email || "");
                            return (
                                <button
                                    key={`${row.source}:${row.email}`}
                                    type="button"
                                    style={{
                                        ...S.participantCard,
                                        ...(isPrimary ? S.participantCardPrimary : {}),
                                        ...(isActive ? S.participantCardActive : {}),
                                    }}
                                    onClick={() => openParticipantDetail(row)}
                                >
                                    <div style={S.participantIdentity}>
                                        <div style={S.participantName}>{row.name}</div>
                                        <div style={S.participantEmail}>{row.email}</div>
                                    </div>
                                    <div style={S.participantCardSide}>
                                        <span style={S.participantBadge}>{row.source.toUpperCase()}</span>
                                        <span style={S.participantChevron}>›</span>
                                    </div>
                                </button>
                            );
                        })}
                    </div>
                ) : (
                    <div style={S.collapsedSummary}>
                        {participants.map((row) => row.name).slice(0, 4).join(" · ")}
                        {participants.length > 4 ? ` +${participants.length - 4}` : ""}
                    </div>
                            )}
                        </div>

                        <div style={S.participantPane}>
                            {!activeParticipant ? (
                                <PanelState
                                    compact
                                    tone="empty"
                                    title="Seleciona um participante"
                                    description="Abre um participante para ver o estado no Odoo, criar ou atualizar o contacto e consultar as ligacoes."
                                />
                            ) : participantDetailLoading ? (
                                <PanelState
                                    compact
                                    tone="loading"
                                    title="A abrir participante"
                                    description={`A carregar contacto e ligacoes de ${activeParticipant.name}.`}
                                />
                            ) : participantDetailError ? (
                                <PanelState compact tone="error" title="Falha ao carregar participante" description={participantDetailError} />
                            ) : (
                                <div style={S.participantDetail}>
                                    <div style={S.participantDetailHead}>
                                        <button
                                            type="button"
                                            style={S.secondaryAction}
                                            onClick={() => {
                                                setActiveParticipantEmail(null);
                                                setActiveParticipantCollection(null);
                                            }}
                                        >
                                            Voltar
                                        </button>
                                        <div style={S.participantDetailTitleWrap}>
                                            <div style={S.participantDetailTitle}>{activeParticipant.name}</div>
                                            <div style={S.participantDetailEmail}>{activeParticipant.email}</div>
                                        </div>
                                        <span style={S.participantBadge}>{activeParticipant.source.toUpperCase()}</span>
                                    </div>

                                    <>
                                    <div style={S.participantDetailMeta}>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Estado</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? "Contacto existente no Odoo" : "Ainda sem contacto no Odoo"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? `ID ${participantDetailPartner.id}${participantCompany ? ` · ${participantCompany}` : ""}`
                                                    : "Podes criar o contacto a partir deste participante."}
                                            </div>
                                        </div>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Storage central</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? `${participantDetailLinks.length} email(s)` : "Sem ligacoes"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? (participantCurrentConversationLinked ? "Este email ja esta ligado ao contacto." : "Ainda nao ha ligacao deste email ao contacto.")
                                                    : "Cria primeiro o contacto para poderes consultar ligacoes."}
                                            </div>
                                        </div>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Relacoes Odoo</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? `${participantRelationTotal} registo(s)` : "Sem relacoes"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? (participantRelationPreview || "Sem relacoes nativas encontradas para este contacto.")
                                                    : "Cria primeiro o contacto para poderes consultar ligacoes nativas no Odoo."}
                                            </div>
                                        </div>
                                    </div>

                                    {participantDetailPartner ? (
                                        <>
                                            <div style={S.participantSummaryGrid}>
                                                {participantDetailPartner.function ? (
                                                    <div style={S.participantSummaryItem}>
                                                        <div style={S.detailKicker}>Funcao</div>
                                                        <div style={S.detailLine}>{participantDetailPartner.function}</div>
                                                    </div>
                                                ) : null}
                                                {participantCompany ? (
                                                    <div style={S.participantSummaryItem}>
                                                        <div style={S.detailKicker}>Empresa</div>
                                                        <div style={S.detailLine}>{participantCompany}</div>
                                                    </div>
                                                ) : null}
                                                {participantDetailPartner.phone ? (
                                                    <div style={S.participantSummaryItem}>
                                                        <div style={S.detailKicker}>Telefone</div>
                                                        <div style={S.detailLine}>{participantDetailPartner.phone}</div>
                                                    </div>
                                                ) : null}
                                                {participantDetailPartner.mobile ? (
                                                    <div style={S.participantSummaryItem}>
                                                        <div style={S.detailKicker}>Telemovel</div>
                                                        <div style={S.detailLine}>{participantDetailPartner.mobile}</div>
                                                    </div>
                                                ) : null}
                                            </div>

                                            <div style={S.participantDetailActions}>
                                                <button
                                                    type="button"
                                                    style={S.primaryAction}
                                                    onClick={() => openParticipantEditor(activeParticipant, Number(participantDetailPartner.id))}
                                                >
                                                    Atualizar
                                                </button>
                                                <button
                                                    type="button"
                                                    style={S.secondaryAction}
                                                    onClick={() => linkParticipantToCurrentEmail(activeParticipant)}
                                                    disabled={participantActionBusy || participantCurrentConversationLinked}
                                                >
                                                    {participantCurrentConversationLinked ? "Ja ligado" : "Ligar ao email"}
                                                </button>
                                                <button
                                                    type="button"
                                                    style={S.secondaryAction}
                                                    onClick={() => openOdooRecord("res.partner", Number(participantDetailPartner.id))}
                                                >
                                                    Odoo
                                                </button>
                                            </div>

                                            <div style={S.participantSectionBlock}>
                                                <div style={S.participantSectionHead}>
                                                    <div style={S.sectionTitle}>Colecoes</div>
                                                    <div style={S.sectionHint}>Cada colecao abre num ecran limpo com pesquisa e ordenacao.</div>
                                                </div>

                                                <div style={S.participantSummaryList}>
                                                    {participantDetailRelations.map((section) => (
                                                        <button
                                                            key={section.key}
                                                            type="button"
                                                            style={S.participantCompactCard}
                                                            onClick={() => openParticipantCollection({ kind: "relation", key: section.key })}
                                                        >
                                                            <div style={S.participantCompactCardHead}>
                                                                <div style={S.participantCompactCardTitle}>{section.label}</div>
                                                                <div style={S.participantCompactCardCount}>{section.total}</div>
                                                            </div>
                                                            <div style={S.participantCompactCardCopy}>{summarizeRelationSection(section)}</div>
                                                        </button>
                                                    ))}

                                                    <button
                                                        type="button"
                                                        style={S.participantCompactCard}
                                                        onClick={() => openParticipantCollection({ kind: "storage" })}
                                                    >
                                                        <div style={S.participantCompactCardHead}>
                                                            <div style={S.participantCompactCardTitle}>Emails ligados</div>
                                                            <div style={S.participantCompactCardCount}>{participantDetailLinks.length}</div>
                                                        </div>
                                                        <div style={S.participantCompactCardCopy}>
                                                            {participantDetailLinks.length
                                                                ? `${participantDetailLinks.length} email(s) no storage central deste contacto.`
                                                                : "Sem emails ligados no storage central."}
                                                        </div>
                                                    </button>
                                                </div>
                                            </div>
                                        </>
                                    ) : (
                                        <div style={S.participantDetailActions}>
                                            <button
                                                type="button"
                                                style={S.primaryAction}
                                                onClick={() => openParticipantEditor(activeParticipant)}
                                            >
                                                Criar contacto
                                            </button>
                                        </div>
                                    )}
                                    </>
                                </div>
                            )}
                        </div>

                        <div style={S.participantPane}>
                            {!activeParticipant ? (
                                <PanelState
                                    compact
                                    tone="empty"
                                    title="Seleciona um participante"
                                    description="Abre primeiro o detalhe do participante para depois escolheres a colecao."
                                />
                            ) : !activeParticipantCollection ? (
                                <PanelState
                                    compact
                                    tone="empty"
                                    title="Escolhe uma colecao"
                                    description="Carrega num card de ligacoes para abrir uma lista limpa, com pesquisa e ordenacao."
                                />
                            ) : (
                                <div style={S.participantDetail}>
                                    <div style={S.participantDetailHead}>
                                        <button
                                            type="button"
                                            style={S.secondaryAction}
                                            onClick={() => setActiveParticipantCollection(null)}
                                        >
                                            Voltar
                                        </button>
                                        <div style={S.participantDetailTitleWrap}>
                                            <div style={S.participantDetailTitle}>{participantCollectionTitle}</div>
                                            <div style={S.participantDetailEmail}>
                                                {activeParticipantCollection?.kind === "relation"
                                                    ? `${activeParticipant.name} · ${activeRelationSection?.total || 0} registo(s)`
                                                    : activeParticipantCollection?.kind === "storage"
                                                        ? `${activeParticipant.name} · ${participantDetailLinks.length} email(s)`
                                                        : activeParticipant.email}
                                            </div>
                                        </div>
                                        <span style={S.participantBadge}>LISTA</span>
                                    </div>

                                    <div style={S.participantDetailMeta}>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Estado</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? "Contacto existente no Odoo" : "Ainda sem contacto no Odoo"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? `ID ${participantDetailPartner.id}${participantCompany ? ` · ${participantCompany}` : ""}`
                                                    : "Podes criar o contacto a partir deste participante."}
                                            </div>
                                        </div>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Storage central</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? `${participantDetailLinks.length} email(s) ligados` : "Sem ligacoes"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? (participantCurrentConversationLinked ? "Este email ja esta ligado ao contacto." : "Ainda nao ha ligacao deste email ao contacto.")
                                                    : "Cria primeiro o contacto para poderes consultar ligacoes."}
                                            </div>
                                        </div>
                                        <div style={S.participantInfoCard}>
                                            <div style={S.detailKicker}>Relacoes Odoo</div>
                                            <div style={S.detailValue}>{participantDetailPartner?.id ? `${participantRelationTotal} registo(s)` : "Sem relacoes"}</div>
                                            <div style={S.detailCopy}>
                                                {participantDetailPartner?.id
                                                    ? (participantRelationPreview || "Sem relacoes nativas encontradas para este contacto.")
                                                    : "Cria primeiro o contacto para poderes consultar ligacoes nativas no Odoo."}
                                            </div>
                                        </div>
                                    </div>

                                    <div style={S.participantCollectionTools}>
                                        <input
                                            type="text"
                                            style={S.participantSearchInput}
                                            value={participantCollectionQuery}
                                            onChange={(event) => setParticipantCollectionQuery(event.target.value)}
                                            placeholder={
                                                activeParticipantCollection.kind === "relation"
                                                    ? "Pesquisar por marca, empresa, cliente, referencia ou texto livre..."
                                                    : "Pesquisar por assunto, remetente ou link..."
                                            }
                                        />
                                        <div style={S.participantSortWrap}>
                                            <label style={S.detailKicker}>Filtrar</label>
                                            <select
                                                style={S.participantSortSelect}
                                                value={participantCollectionFilter}
                                                onChange={(event) => setParticipantCollectionFilter(event.target.value as ParticipantCollectionFilter)}
                                            >
                                                <option value="all">Tudo</option>
                                                <option value="title">{activeParticipantCollection.kind === "relation" ? "Titulo" : "Assunto"}</option>
                                                <option value="meta">{activeParticipantCollection.kind === "relation" ? "Contexto" : "Remetente"}</option>
                                                <option value="detail">{activeParticipantCollection.kind === "relation" ? "Detalhe" : "Link / IDs"}</option>
                                            </select>
                                        </div>
                                        <div style={S.participantSortWrap}>
                                            <label style={S.detailKicker}>Ordenar</label>
                                            <select
                                                style={S.participantSortSelect}
                                                value={participantCollectionSort}
                                                onChange={(event) => setParticipantCollectionSort(event.target.value as ParticipantCollectionSort)}
                                            >
                                                {activeParticipantCollection.kind === "storage" ? <option value="recent">Mais recente</option> : null}
                                                <option value="title_asc">A-Z</option>
                                                <option value="title_desc">Z-A</option>
                                            </select>
                                        </div>
                                    </div>

                                    {activeParticipantCollection.kind === "relation" ? (
                                        !activeRelationSection || !filteredRelationItems.length ? (
                                            <PanelState
                                                compact
                                                tone="empty"
                                                title="Sem resultados"
                                                description={activeRelationSection ? "A pesquisa nao devolveu registos nesta colecao." : "Esta colecao deixou de estar disponivel."}
                                            />
                                        ) : (
                                            <div style={S.collectionList}>
                                                {filteredRelationItems.map((item) => (
                                                    <div key={`${item.model}:${item.recordId}`} style={S.collectionRow}>
                                                        <div style={S.collectionRowMain}>
                                                            <div style={S.collectionRowTitle}>{item.title || `#${item.recordId}`}</div>
                                                            {item.meta ? <div style={S.collectionRowMeta}>{item.meta}</div> : null}
                                                            {item.secondary ? <div style={S.collectionRowMeta}>{item.secondary}</div> : null}
                                                        </div>
                                                        <div style={S.collectionRowActions}>
                                                            {canOpenCrm2Editor(item.model) ? (
                                                                <button
                                                                    type="button"
                                                                    style={S.linkAction}
                                                                    onClick={() => openDialog("edit", { model: item.model, recordId: String(item.recordId) })}
                                                                >
                                                                    Editor v2
                                                                </button>
                                                            ) : null}
                                                            <button
                                                                type="button"
                                                                style={canOpenCrm2Editor(item.model) ? S.linkActionMuted : S.linkAction}
                                                                onClick={() => openOdooRecord(item.model, item.recordId)}
                                                            >
                                                                Odoo
                                                            </button>
                                                        </div>
                                                    </div>
                                                ))}
                                            </div>
                                        )
                                    ) : !filteredStorageLinks.length ? (
                                        <PanelState
                                            compact
                                            tone="empty"
                                            title="Sem resultados"
                                            description="Nao encontrei emails ligados com os filtros atuais."
                                        />
                                    ) : (
                                        <div style={S.collectionList}>
                                            {filteredStorageLinks.map((entry, index) => (
                                                <div key={`${entry.conversationId || entry.itemId || entry.subject || "link"}:${index}`} style={S.collectionRow}>
                                                    <div style={S.collectionRowMain}>
                                                        <div style={S.collectionRowTitle}>{entry.subject || "Email sem assunto"}</div>
                                                        <div style={S.collectionRowMeta}>
                                                            {`Ligacao ${index + 1} - ${formatLinkMoment(entry)}`}
                                                            {entry.fromEmail ? ` - ${entry.fromEmail}` : ""}
                                                        </div>
                                                    </div>
                                                    <div style={S.collectionRowActions}>
                                                        {(entry.emailWebLink || entry.url) ? (
                                                            <button
                                                                type="button"
                                                                style={S.linkAction}
                                                                onClick={() => window.open(entry.emailWebLink || entry.url, "_blank")}
                                                            >
                                                                Abrir email
                                                            </button>
                                                        ) : null}
                                                        <button
                                                            type="button"
                                                            style={S.linkActionMuted}
                                                            onClick={() => openOdooRecord("res.partner", Number(participantDetailPartner?.id || 0))}
                                                        >
                                                            Odoo
                                                        </button>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    )}
                                </div>
                            )}
                        </div>
                    </div>
                </div>
            </section>

            {!activeParticipant ? (
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
            ) : null}

            {!activeParticipant ? (
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
            ) : null}
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
    participantViewport: {
        overflow: "hidden",
    },
    participantTrack: {
        display: "flex",
        width: "300%",
        transition: "transform 220ms ease",
    },
    participantPane: {
        width: "33.3333%",
        minWidth: 0,
        paddingRight: 8,
        boxSizing: "border-box",
    },
    participantList: {
        display: "grid",
        gap: 8,
    },
    participantCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "6px 10px",
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        gap: 10,
        width: "100%",
        textAlign: "left",
        cursor: "pointer",
    },
    participantCardPrimary: {
        borderColor: "#0C66E4",
        background: "#EFF6FF",
    },
    participantCardActive: {
        boxShadow: "0 0 0 2px rgba(12,102,228,0.16) inset",
    },
    participantIdentity: {
        minWidth: 0,
        flex: 1,
    },
    participantCardSide: {
        display: "grid",
        justifyItems: "end",
        gap: 6,
        flexShrink: 0,
    },
    participantName: {
        fontSize: 11,
        fontWeight: 700,
        color: "#172B4D",
    },
    participantEmail: {
        fontSize: 10,
        color: "#42526E",
    },
    participantBadge: {
        display: "inline-flex",
        alignItems: "center",
        justifyContent: "center",
        minWidth: 32,
        borderRadius: 999,
        padding: "2px 7px",
        fontSize: 9,
        fontWeight: 800,
        background: "#DFE1E6",
        color: "#42526E",
    },
    participantChevron: {
        fontSize: 12,
        fontWeight: 800,
        color: "#6B778C",
    },
    participantDetail: {
        display: "grid",
        gap: 10,
    },
    participantDetailHead: {
        display: "grid",
        gridTemplateColumns: "auto minmax(0, 1fr) auto",
        gap: 10,
        alignItems: "center",
    },
    participantDetailTitleWrap: {
        minWidth: 0,
    },
    participantDetailTitle: {
        fontSize: 14,
        fontWeight: 800,
        color: "#172B4D",
    },
    participantDetailEmail: {
        marginTop: 2,
        fontSize: 11,
        color: "#42526E",
        wordBreak: "break-all",
    },
    participantDetailMeta: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(160px, 1fr))",
        gap: 8,
    },
    participantInfoCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "grid",
        gap: 4,
    },
    detailKicker: {
        fontSize: 10,
        fontWeight: 800,
        color: "#6B778C",
        textTransform: "uppercase",
        letterSpacing: "0.05em",
    },
    detailValue: {
        fontSize: 13,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.3,
    },
    detailCopy: {
        fontSize: 11,
        color: "#5E6C84",
        lineHeight: 1.45,
    },
    participantSummaryGrid: {
        display: "grid",
        gridTemplateColumns: "repeat(auto-fit, minmax(140px, 1fr))",
        gap: 8,
    },
    participantSummaryItem: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#FFFFFF",
        padding: "10px 12px",
        display: "grid",
        gap: 4,
    },
    detailLine: {
        fontSize: 12,
        fontWeight: 600,
        color: "#172B4D",
        lineHeight: 1.35,
    },
    participantDetailActions: {
        display: "flex",
        gap: 8,
        flexWrap: "wrap",
    },
    participantSectionBlock: {
        display: "grid",
        gap: 8,
    },
    participantSectionHead: {
        display: "grid",
        gap: 2,
    },
    participantSummaryList: {
        display: "grid",
        gap: 8,
    },
    participantCompactCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 8,
        background: "#F7F8FA",
        padding: "6px 8px",
        display: "grid",
        gap: 2,
        textAlign: "left",
        cursor: "pointer",
    },
    participantCompactCardHead: {
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 8,
    },
    participantCompactCardTitle: {
        fontSize: 11,
        fontWeight: 800,
        color: "#172B4D",
    },
    participantCompactCardCount: {
        minWidth: 22,
        padding: "2px 6px",
        borderRadius: 999,
        background: "#E9F2FF",
        color: "#0C66E4",
        fontSize: 10,
        fontWeight: 800,
        textAlign: "center",
    },
    participantCompactCardCopy: {
        fontSize: 10,
        lineHeight: 1.25,
        color: "#5E6C84",
        display: "-webkit-box",
        WebkitLineClamp: 1,
        WebkitBoxOrient: "vertical",
        overflow: "hidden",
    },
    participantCollectionTools: {
        display: "grid",
        gridTemplateColumns: "minmax(0, 1fr) repeat(2, minmax(108px, auto))",
        gap: 8,
        alignItems: "end",
    },
    collectionList: {
        display: "grid",
        gap: 0,
        borderTop: "1px solid #DFE1E6",
    },
    collectionRow: {
        display: "grid",
        gridTemplateColumns: "minmax(0, 1fr) auto",
        gap: 10,
        alignItems: "center",
        padding: "10px 0",
        borderBottom: "1px solid #DFE1E6",
    },
    collectionRowMain: {
        minWidth: 0,
        display: "grid",
        gap: 4,
    },
    collectionRowTitle: {
        fontSize: 12,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.3,
    },
    collectionRowMeta: {
        fontSize: 10,
        color: "#5E6C84",
        lineHeight: 1.35,
    },
    collectionRowActions: {
        display: "flex",
        gap: 8,
        flexWrap: "wrap",
        justifyContent: "flex-end",
    },
    participantSearchInput: {
        border: "1px solid #DFE1E6",
        borderRadius: 8,
        background: "#FFFFFF",
        color: "#172B4D",
        padding: "8px 10px",
        fontSize: 11,
        outline: "none",
        width: "100%",
        boxSizing: "border-box",
    },
    participantSortWrap: {
        display: "grid",
        gap: 4,
        minWidth: 108,
    },
    participantSortSelect: {
        border: "1px solid #DFE1E6",
        borderRadius: 8,
        background: "#FFFFFF",
        color: "#172B4D",
        padding: "8px 10px",
        fontSize: 11,
        outline: "none",
    },
    participantRelationGroups: {
        display: "grid",
        gap: 10,
    },
    participantRelationGroupCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#FFFFFF",
        padding: "10px 12px",
        display: "grid",
        gap: 8,
    },
    participantRelationGroupHead: {
        display: "flex",
        alignItems: "center",
        justifyContent: "space-between",
        gap: 8,
    },
    participantLinksList: {
        display: "grid",
        gap: 8,
    },
    participantLinkCard: {
        border: "1px solid #DFE1E6",
        borderRadius: 10,
        background: "#F7F8FA",
        padding: "10px 12px",
        display: "grid",
        gap: 6,
    },
    participantLinkTitle: {
        fontSize: 12,
        fontWeight: 700,
        color: "#172B4D",
        lineHeight: 1.35,
    },
    participantLinkMeta: {
        display: "flex",
        flexWrap: "wrap",
        gap: 8,
        fontSize: 10,
        color: "#5E6C84",
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
