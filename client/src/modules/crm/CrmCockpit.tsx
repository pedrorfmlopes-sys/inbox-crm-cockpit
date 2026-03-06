import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { openCockpitDialog, getEmailBodyText } from "../../office";
import * as Icons from "../../ui/icons";
import { ContactInsight } from "./ContactInsight";
import { OdooCardSkeleton, Skeleton } from "../../ui/SkeletonLoader";
import { aiExtractAnchors, aiGenerate, getOdooAutoLoginUrl, searchOdoo } from "../../api";
import { scanForProtection, MatchResult } from "./triangulationService";
import { ProtectionBanner } from "../../ui/ProtectionBanner";


/**
 * VerticalActionCascade: Ultra-compact glossy pill menu.
 * Strictly 94x26px, 16px radius, glossy effect.
 */
const VerticalActionCascade: React.FC<{ onSelect: (type: string) => void }> = ({ onSelect }) => {
    const [isOpen, setIsOpen] = useState(false);
    const [hoveredBtn, setHoveredBtn] = useState<string | null>(null);
    const ref = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (ref.current && !ref.current.contains(e.target as Node)) setIsOpen(false);
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    const items = [
        { label: "Tarefa", type: "project.task", icon: "📝" },
        { label: "Ticket", type: "helpdesk.ticket", icon: "🎫" },
        { label: "Lead", type: "crm.lead", icon: "🎯" },
        { label: "Projeto", type: "project.project", icon: "🏗️" },
        { label: "Contato", type: "res.partner", icon: "👤" },
    ];

    const primaryStyle: React.CSSProperties = {
        ...S.primaryBtn,
        transition: "all 0.18s ease",
        ...(hoveredBtn === "criar" ? {
            background: "linear-gradient(180deg, rgba(110,180,255,0.98) 0%, rgba(20,120,230,0.95) 100%)",
            boxShadow: "0 6px 14px rgba(0,100,210,0.5), inset 0 1px 0 rgba(255,255,255,0.7), inset 0 -1px 0 rgba(0,0,0,0.15)",
            transform: "translateY(-1px)",
        } : {}),
    };

    const secondaryStyle = (key: string): React.CSSProperties => ({
        ...S.secondaryBtn,
        transition: "all 0.18s ease",
        ...(hoveredBtn === key ? {
            background: "linear-gradient(180deg, rgba(230,240,255,0.98) 0%, rgba(195,215,248,0.95) 100%)",
            boxShadow: "0 6px 14px rgba(0,80,200,0.15), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
            transform: "translateY(-1px)",
        } : {}),
    });

    return (
        <div ref={ref} style={{ display: "flex", flexDirection: "column", gap: "4px", width: "94px" }}>
            <button
                style={primaryStyle}
                onClick={() => setIsOpen(!isOpen)}
                onMouseEnter={() => setHoveredBtn("criar")}
                onMouseLeave={() => setHoveredBtn(null)}
                title="Criar Item"
            >
                <Icons.Plus size={11} />
                CRIAR
            </button>

            {isOpen && (
                <div style={{ display: "flex", flexDirection: "column", gap: "4px" }}>
                    {items.map(item => (
                        <button
                            key={item.type}
                            style={secondaryStyle(item.type)}
                            title={item.label}
                            onMouseEnter={() => setHoveredBtn(item.type)}
                            onMouseLeave={() => setHoveredBtn(null)}
                            onClick={() => {
                                onSelect(item.type);
                                setIsOpen(false);
                            }}
                        >
                            <span style={{ fontSize: "11px", lineHeight: 1, flexShrink: 0 }}>{item.icon}</span>
                            {item.label.toUpperCase()}
                        </button>
                    ))}
                </div>
            )}
        </div>
    );
};

export const CrmCockpit: React.FC = () => {
    const { ctx, bodyText, attachments, meta, links, msg, refreshLinks, setMsg, isLoading: isContextLoading, settings } = useCockpit() as any;

    const customModels = settings ? {
        openaiModelFast: settings.openaiModelFast,
        openaiModelQuality: settings.openaiModelQuality,
        geminiModel: settings.geminiModel,
        openaiApiKey: settings.openaiApiKey,
        geminiApiKey: settings.geminiApiKey,
    } : {};

    const [isLinkedExpanded, setIsLinkedExpanded] = useState(true);
    const [isAnchorsLoading, setIsAnchorsLoading] = useState(false);
    const [anchors, setAnchors] = useState<any>(null);
    const [contact, setContact] = useState<any>(null);
    const [isContactLoading, setIsContactLoading] = useState(false);
    const [protection, setProtection] = useState<MatchResult | null>(null);
    const [isDrafting, setIsDrafting] = useState(false);
    const [isBriefingLoading, setIsBriefingLoading] = useState(false);
    const [briefing, setBriefing] = useState<string | null>(null);
    const [voiceCommand, setVoiceCommand] = useState("");
    const [isVoiceLoading, setIsVoiceLoading] = useState(false);
    const [isEmailInsightsLoading, setIsEmailInsightsLoading] = useState(false);
    const [emailInsights, setEmailInsights] = useState<Array<{ model: string; recordId: number; recordName: string }>>([]);

    async function handleGetBriefing() {
        setIsBriefingLoading(true);
        try {
            const { get30SecondBriefing } = await import("./HistorySummaryService");
            const b = await get30SecondBriefing({
                outlookHistory: "Recent collaboration on Project X.", // In real app, fetch from Outlook
                odooChatter: meta?.chatter || "Client interested in premium finishes.",
                protectionStatus: protection?.isProtected ? `PROTECTED (${protection.matchedProject?.projectName})` : "FREE"
            }, customModels);
            setBriefing(b);
        } catch (e) {
            console.error("Briefing failed:", e);
        } finally {
            setIsBriefingLoading(false);
        }
    }

    async function handleVoiceAction() {
        if (!voiceCommand.trim()) return;
        setIsVoiceLoading(true);
        try {
            const { aiVoiceCommand } = await import("../../api");
            const res = await aiVoiceCommand(voiceCommand, { anchors, protection, customModels });
            if (res.ok) {
                setVoiceCommand("");
                // EXECUTE CHAINED ACTIONS
                for (const action of res.actions) {
                    if (action === "GENERATE_BRIEFING") await handleGetBriefing();
                    if (action === "EXTRACT_ANCHORS") await handleScanAnchors();
                    if (action === "DRAFT_REJECTION") await handleDraftRejection();
                }
            }
        } catch (e) {
            console.error("Voice command failed:", e);
        } finally {
            setIsVoiceLoading(false);
        }
    }

    // Initial check: Extract anchors when email changes
    useEffect(() => {
        setBriefing(null);   // Reset briefing so old email data doesn't linger
        setAnchors(null);    // Reset anchors
        setContact(null);    // Reset contact
        setProtection(null); // Reset protection
        setEmailInsights([]);
        if (ctx.conversationId) {
            handleScanAnchors();
            loadContact();
            loadEmailInsights();
        }
    }, [ctx.conversationId]);

    async function loadEmailInsights() {
        const email = String(ctx.fromEmail || "").trim();
        if (!email) return;

        setIsEmailInsightsLoading(true);
        try {
            const partners: any[] = await searchOdoo({
                model: "res.partner",
                domain: [["email", "ilike", email]],
                fields: ["id", "name", "display_name", "email"],
                limit: 10,
            });
            const partnerIds = partners.map((p: any) => Number(p.id)).filter(Boolean);

            const leadsByEmail: any[] = await searchOdoo({
                model: "crm.lead",
                domain: [["email_from", "ilike", email]],
                fields: ["id", "name", "display_name", "partner_id"],
                limit: 10,
            });

            const leadsByPartner: any[] = partnerIds.length
                ? await searchOdoo({
                    model: "crm.lead",
                    domain: [["partner_id", "in", partnerIds]],
                    fields: ["id", "name", "display_name", "partner_id"],
                    limit: 20,
                })
                : [];

            const allLeads = [...(leadsByEmail || []), ...(leadsByPartner || [])];
            const leadIds = Array.from(new Set(allLeads.map((l: any) => Number(l.id)).filter(Boolean)));

            const projects: any[] = partnerIds.length
                ? await searchOdoo({
                    model: "project.project",
                    domain: [["partner_id", "in", partnerIds]],
                    fields: ["id", "name", "display_name", "partner_id"],
                    limit: 20,
                })
                : [];

            const tickets: any[] = partnerIds.length
                ? await searchOdoo({
                    model: "helpdesk.ticket",
                    domain: [["partner_id", "in", partnerIds]],
                    fields: ["id", "name", "display_name", "partner_id", "stage_id", "priority"],
                    limit: 20,
                })
                : [];

            const tasks: any[] = leadIds.length
                ? await searchOdoo({
                    model: "project.task",
                    domain: [["lead_id", "in", leadIds]],
                    fields: ["id", "name", "display_name", "lead_id", "project_id", "stage_id"],
                    limit: 20,
                })
                : [];

            const all = [
                ...(partners || []).map((r: any) => ({ model: "res.partner", recordId: Number(r.id), recordName: r.display_name || r.name || `#${r.id}` })),
                ...(allLeads || []).map((r: any) => ({ model: "crm.lead", recordId: Number(r.id), recordName: r.display_name || r.name || `#${r.id}` })),
                ...(projects || []).map((r: any) => ({ model: "project.project", recordId: Number(r.id), recordName: r.display_name || r.name || `#${r.id}` })),
                ...(tasks || []).map((r: any) => ({ model: "project.task", recordId: Number(r.id), recordName: r.display_name || r.name || `#${r.id}` })),
                ...(tickets || []).map((r: any) => ({ model: "helpdesk.ticket", recordId: Number(r.id), recordName: r.display_name || r.name || `#${r.id}` })),
            ].filter((x) => x.recordId);

            const dedup = new Map<string, { model: string; recordId: number; recordName: string }>();
            for (const row of all) dedup.set(`${row.model}:${row.recordId}`, row);
            setEmailInsights(Array.from(dedup.values()));
        } catch (e) {
            console.error("[crm] Failed to load email insights", e);
            setEmailInsights([]);
        } finally {
            setIsEmailInsightsLoading(false);
        }
    }

    async function handleScanAnchors() {
        setIsAnchorsLoading(true);
        setProtection(null); // Reset
        try {
            const body = await getEmailBodyText();
            if (body) {
                // START PARALLEL SCAN: Local-first triangulation starts
                // Walkthrough Update:
                // - [x] **Odoo Interactivity**:
                //     - **Smart Alerts**: The "Odoo" action button in the CRM tab now informs the user if the connection is missing, rather than failing silently.
                //     - **Deep Links**: Odoo status footers in all tabs now function as direct links to the configured Odoo instance.
                //     - **Status Sync**: Synchronized the "Green Dot" indicator with actual Odoo metadata availability to prevent misleading connection statuses.
                // Note: Real triangulation needs anchors. We do a partial scan with body/subject first
                // OR we wait for Flash Anchors (sub-second) and then scan IndexedDB instantly.
                // FLASH is sub-second, satisfying the < 0.8s goal.
                const emailContext = {
                    subject: ctx.subject || "",
                    from: { name: ctx.fromName || "", email: ctx.fromEmail || "" },
                    to: ctx.toRecipients || [],
                    cc: ctx.ccRecipients || [],
                    bodyText: body
                };
                const res = await aiExtractAnchors(body, customModels, emailContext);
                if (res.ok) {
                    setAnchors(res.anchors);
                    // LOCAL SCAN IS INSTANT (<10ms)
                    const p = await scanForProtection(res.anchors);
                    setProtection(p);
                }
            }
        } catch (e) {
            console.error("[crm] Anchor scan failed:", e);
        } finally {
            setIsAnchorsLoading(false);
        }
    }

    
    async function loadContact() {
        const email = (ctx.fromEmail || "").trim();
        if (!email) return;
        setIsContactLoading(true);
        try {
            // Exact match first
            const recs: any[] = await searchOdoo({
                model: "res.partner",
                domain: [["email", "=", email]],
                fields: ["id", "name", "email", "phone", "mobile", "function", "company_name", "parent_id"],
                limit: 1
            });
            const r = Array.isArray(recs) && recs.length ? recs[0] : null;

            if (r) {
                const company =
                    (Array.isArray(r.parent_id) ? r.parent_id[1] : null) ||
                    r.company_name ||
                    "";
                setContact({
                    id: r.id,
                    name: r.name || ctx.fromName || "Desconhecido",
                    email: r.email || email,
                    phone: r.phone,
                    mobile: r.mobile,
                    role: r.function || undefined,
                    company: company || undefined,
                });
            } else {
                // fallback: keep minimal
                setContact({
                    id: null,
                    name: ctx.fromName || "Desconhecido",
                    email,
                });
            }
        } catch (e) {
            console.error("[crm] Failed to load contact from Odoo", e);
            setContact({
                id: null,
                name: ctx.fromName || "Desconhecido",
                email,
            });
        } finally {
            setIsContactLoading(false);
        }
    }

async function handleDraftRejection() {
        if (!protection?.matchedProject) return;
        setIsDrafting(true);
        try {
            const res = await aiGenerate({
                action: "reply",
                inputText: `O projeto "${protection.matchedProject.projectName}" já está protegido para o distribuidor "${protection.matchedProject.distributor}". Redige um email diplomático a explicar que não podemos cotar diretamente ou para outro canal.`,
                mode: "quality", // Use Pro for the "Second Brain" quality
                customModels,
            });
            if (res.ok && res.text) {
                // In real app, this would insert into Outlook draft
                alert("Rascunho Diplomático Gerado!");
            }
        } catch (e: any) {
            setMsg(e.message);
        } finally {
            setIsDrafting(false);
        }
    }

    async function openDialog(targetMode: "new" | "add" | "edit", extra?: Record<string, string>) {
        if (!ctx.conversationId && targetMode !== "edit") {
            setMsg("Seleciona um email primeiro.");
            return;
        }

        // Pass large data via localStorage to avoid URL length limits
        try {
            localStorage.setItem("ic_bridge_body", bodyText || "");
            localStorage.setItem("ic_bridge_atts", JSON.stringify(attachments || []));
        } catch (e) {
            console.error("[crm] Failed to save transition data", e);
        }

        try {
            await openCockpitDialog({
                mode: targetMode,
                conversationId: ctx.conversationId || "",
                internetMessageId: ctx.internetMessageId || "",
                subject: ctx.subject || "",
                fromEmail: ctx.fromEmail || "",
                fromName: ctx.fromName || "",
                receivedAtIso: ctx.receivedDateTimeIso || "",
                toR: ctx.toRecipients || [],
                ccR: ctx.ccRecipients || [],
                ...(extra || {}),
            });
            await refreshLinks();
        } catch (e: any) {
            setMsg(e?.message ?? String(e));
        }
    }

    //## Sprint 10: Connectivity Logic Sync
    // - [x] Align Navigation Dot with Odoo Metadata availability
    // - [x] Proactive Metadata refresh in connectivity heartbeat
    // - [x] Synchronize Odoo status footer with actual Odoo metadata availability
    // Contact data (from Odoo when possible)
    const displayContact = {
        name: contact?.name || ctx.fromName || "Desconhecido",
        email: contact?.email || ctx.fromEmail || "",
        role: contact?.role || undefined,
        company: contact?.company || undefined,
    };


    const isInitialLoading = isContextLoading || (links.length === 0 && !meta);

    return (
        <div style={S.container}>
            {/* HubSpot-style Sticky Header */}
            <ContactInsight
                contact={displayContact}
                onViewInOdoo={() => {
                    const baseUrl = meta?.baseUrl || (settings as any)?.odooUrl;
                    const db = (settings as any)?.odooDb || meta?.db || "divitek";
                    const id = contact?.id;
                    const target = id ? `/web?db=${encodeURIComponent(db)}#id=${encodeURIComponent(String(id))}&model=res.partner&view_type=form` : `/web?db=${encodeURIComponent(db)}`;
                    window.open(getOdooAutoLoginUrl((settings as any)?.odooSessionToken || null, target, baseUrl), "_blank");
                }}
            />

            <div style={S.scrollArea}>
                {/* Fingerprint / Anchors Area (Visualized if loading) */}
                {isAnchorsLoading && (
                    <div style={{ marginBottom: "12px", border: "1px dashed #dbeafe", padding: "8px", borderRadius: "6px" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px", marginBottom: "4px" }}>
                            <Icons.Sparkles size={10} color="#2563eb" />
                            <span style={{ fontSize: "9px", fontWeight: 700, textTransform: "uppercase", color: "#2563eb" }}>Extraindo Âncoras...</span>
                        </div>
                        <Skeleton width="60%" height="10px" marginBottom="4px" />
                        <Skeleton width="40%" height="10px" />
                    </div>
                )}

                {/* THE MOAT: Protection Banner */}
                {protection?.isProtected && protection.matchedProject && (
                    <ProtectionBanner
                        project={protection.matchedProject}
                        confidence={protection.confidence}
                        reason={protection.reason}
                        onDraftRejection={handleDraftRejection}
                        isDrafting={isDrafting}
                    />
                )}

                {/* 30-Second Briefing Area */}
                <div style={{ ...S.section, padding: "10px", borderColor: "#bfdbfe", background: "#eff6ff" }}>
                    <div style={{ display: "flex", justifyContent: "space-between", alignItems: "center", marginBottom: isBriefingLoading || briefing ? "8px" : "0" }}>
                        <div style={{ fontSize: "10px", fontWeight: 800, color: "#2563eb", textTransform: "uppercase" }}>30-Sec Briefing</div>
                        <button
                            onClick={handleGetBriefing}
                            disabled={isBriefingLoading}
                            style={{ background: "none", border: "none", color: "#2563eb", cursor: "pointer", fontSize: "10px", fontWeight: 700, display: "flex", alignItems: "center", gap: "4px" }}
                            title="Gerar ou Atualizar Resumo"
                        >
                            {isBriefingLoading ? <Icons.RefreshCw size={10} className="animate-spin" /> : <Icons.Sparkles size={10} />}
                            {briefing ? "Actualizar" : "Gerar Resumo"}
                        </button>
                    </div>
                    {briefing && (
                        <div style={{ fontSize: "11px", color: "#1e3a8a", lineHeight: "1.4", whiteSpace: "pre-wrap" }}>
                            {briefing}
                        </div>
                    )}
                </div>

                <div style={S.actionRow}>
                    <VerticalActionCascade
                        onSelect={(type) => {
                            if (type === "res.partner") openDialog("new", { model: "res.partner" });
                            else if (type === "project.task") openDialog("new", { model: "project.task" });
                            else if (type === "helpdesk.ticket") openDialog("new", { model: "helpdesk.ticket" });
                            else if (type === "crm.lead") openDialog("new", { model: "crm.lead" });
                            else if (type === "project.project") openDialog("new", { model: "project.project" });
                            else if (type === "res.partner") openDialog("new", { model: "res.partner" });
                        }}
                    />
                    <button style={S.secondaryBtn} onClick={() => openDialog("add")}>
                        <Icons.Link size={12} />
                        LIGAR
                    </button>
                </div>

                {msg && <div style={S.alert}>{msg}</div>}

                <div style={S.section}>
                    <div style={S.sectionHeader}>
                        <h3 style={S.sectionTitle}>Encontrados no Odoo por este email ({emailInsights.length})</h3>
                    </div>
                    <div style={S.accordionContent}>
                        {isEmailInsightsLoading ? (
                            <>
                                <OdooCardSkeleton />
                                <OdooCardSkeleton />
                            </>
                        ) : !emailInsights.length ? (
                            <div style={S.emptyState}>
                                <p>Nenhum contacto/lead/projeto/tarefa/ticket encontrado por email.</p>
                            </div>
                        ) : (
                            <div style={S.cardList}>
                                {emailInsights.map((link) => (
                                    <OdooCard
                                        key={`insight-${link.model}-${link.recordId}`}
                                        link={link}
                                        meta={meta}
                                        settings={settings}
                                        onEdit={() => openDialog("edit", { model: link.model, recordId: String(link.recordId) })}
                                    />
                                ))}
                            </div>
                        )}
                    </div>
                </div>

                {/* Jira-style Collapsible Bucket */}
                <div style={S.section}>
                    <div style={S.sectionHeader} onClick={() => setIsLinkedExpanded(!isLinkedExpanded)}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <div style={{ transform: isLinkedExpanded ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform 0.2s' }}>
                                <Icons.ExternalLink size={12} style={{ transform: 'rotate(-45deg)' }} />
                            </div>
                            <h3 style={S.sectionTitle}>Ligados a esta conversa ({links.length})</h3>
                        </div>
                        <button style={S.refreshBtn} onClick={(e) => { e.stopPropagation(); refreshLinks(); }}>
                            <Icons.RefreshCw size={12} />
                        </button>
                    </div>

                    {isLinkedExpanded && (
                        <div style={S.accordionContent}>
                            {isInitialLoading ? (
                                <>
                                    <OdooCardSkeleton />
                                    <OdooCardSkeleton />
                                </>
                            ) : !links.length ? (
                                <div style={S.emptyState}>
                                    <p>{!meta ? "Odoo não configurado." : "Nenhum registo associado."}</p>
                                </div>
                            ) : (
                                <div style={S.cardList}>
                                    {links.map((link: any) => (
                                        <OdooCard
                                            key={`${link.model}-${link.recordId}`}
                                            link={link}
                                            meta={meta}
                                            settings={settings}
                                            onEdit={() => openDialog("edit", { model: link.model, recordId: String(link.recordId) })}
                                        />
                                    ))}
                                </div>
                            )}
                        </div>
                    )}
                </div>
            </div>

            <div style={S.voiceBar}>
                <div style={S.voiceInputWrapper}>
                    <Icons.MessageSquare size={14} style={{ opacity: 0.5 }} />
                    <input
                        type="text"
                        placeholder="Comando de Voz / IA (ex: 'Analisa e rascunha')"
                        style={S.voiceInput}
                        value={voiceCommand}
                        onChange={(e) => setVoiceCommand(e.target.value)}
                        onKeyDown={(e) => e.key === 'Enter' && handleVoiceAction()}
                    />
                    {isVoiceLoading ? (
                        <Icons.RefreshCw size={14} className="animate-spin" />
                    ) : (
                        <button
                            onClick={handleVoiceAction}
                            style={S.voiceSendBtn}
                            title="Enviar Comando"
                        >
                            <Icons.ArrowRight size={14} />
                        </button>
                    )}
                </div>
            </div>

            {meta ? (
                <a
                    href={getOdooAutoLoginUrl(settings?.odooSessionToken || null, `/web?db=${encodeURIComponent((settings as any)?.odooDb || meta?.db || "divitek")}`, meta.baseUrl)}
                    target="_blank"
                    rel="noreferrer"
                    style={{ ...S.footer, textDecoration: 'none', cursor: 'pointer' }}
                    title="Abrir Odoo (Forced DB)"
                >
                    {String(meta?.baseUrl || "").replace(/https?:\/\//, '')} • DIVITEK
                </a>
            ) : (
                <div style={S.footer} onClick={() => setMsg("Odoo não configurado.")}>
                    Odoo: Desconectado
                </div>
            )}
        </div>
    );
};

const OdooCard: React.FC<{ link: any; meta: any; settings: any; onEdit: () => void }> = ({ link, meta, settings, onEdit }) => {
    const modelPrefix = link.model.split(".")[0]?.toUpperCase() || link.model.toUpperCase();
    const db = "divitek"; // Strictly forced as per Sprint 14 requirements
    const target = `/web?db=${encodeURIComponent(db)}#id=${link.recordId}&model=${encodeURIComponent(link.model)}&view_type=form`;
    const url = getOdooAutoLoginUrl(settings?.odooSessionToken || null, target, meta?.baseUrl);

    const getStatusInfo = (model: string) => {
        // Jira Lozenge Styles
        if (model.includes("project")) return {
            bg: "#DEEBFF",
            color: "#0747A6",
            label: "PROJECT",
            border: "none"
        };
        if (model.includes("lead")) return {
            bg: "#FFF0B3",
            color: "#172B4D",
            label: "PROTECTED",
            border: "1px solid #FFC400"
        };
        return {
            bg: "#E3FCEF",
            color: "#006644",
            label: "WON",
            border: "none"
        };
    };

    const status = getStatusInfo(link.model);

    return (
        <div style={S.card}>
            <div style={S.cardHeader}>
                <span style={{
                    ...S.modelTag,
                    background: status.bg,
                    color: status.color,
                    border: status.border
                }}>{status.label}</span>
                <span style={S.recordId}>#{link.recordId}</span>
                {link.priority && (
                    <div style={{ marginLeft: 4, display: "flex", alignItems: "center" }} title={`Priority: ${link.priority}`}>
                        {link.priority === "high" || link.priority === "critical" ? (
                            <Icons.ArrowUp size={14} color="#FF5630" />
                        ) : link.priority === "low" ? (
                            <Icons.ArrowDown size={14} color="#0052CC" />
                        ) : null}
                    </div>
                )}
                <div style={{ flex: 1 }} />
                <button style={S.quickActionBtn} onClick={onEdit} title="Quick Edit">
                    <Icons.Edit size={10} />
                </button>
            </div>

            <div style={S.cardTitle}>{link.recordName || "Sem título"}</div>

            <div style={S.cardFooter}>
                <a href={url} target="_blank" rel="noreferrer" style={S.linkBtn}>
                    <Icons.ExternalLink size={10} />
                    Odoo
                </a>
            </div>
        </div>
    );
};

const S: Record<string, React.CSSProperties> = {
    container: {
        display: "flex",
        flexDirection: "column",
        height: "100vh",
        background: "#FFFFFF",
    },
    scrollArea: {
        flex: 1,
        overflowY: "auto",
        padding: "12px",
        display: "flex",
        flexDirection: "column",
        gap: "12px",
    },
    actionRow: {
        display: "flex",
        gap: "8px",
        padding: "0 4px",
        overflow: "visible"
    },
    primaryBtn: {
        boxSizing: "border-box",
        width: "94px", minWidth: "94px", maxWidth: "94px",
        height: "26px", minHeight: "26px", maxHeight: "26px",
        borderRadius: "16px",
        border: "1px solid rgba(0, 80, 180, 0.4)",
        backdropFilter: "blur(12px)",
        WebkitBackdropFilter: "blur(12px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "5px",
        padding: "0 8px",
        fontSize: "10px",
        fontWeight: 800,
        lineHeight: 1,
        textTransform: "uppercase",
        cursor: "pointer",
        flexShrink: 0,
        margin: 0,
        outline: "none",
        /* Azure Gel 3D */
        background: "linear-gradient(180deg, rgba(80, 160, 255, 0.95) 0%, rgba(0, 100, 210, 0.85) 100%)",
        color: "#FFFFFF",
        boxShadow: "0 4px 10px rgba(0,100,210,0.35), inset 0 1px 0 rgba(255,255,255,0.55), inset 0 -1px 0 rgba(0,0,0,0.15)",
    },
    secondaryBtn: {
        boxSizing: "border-box",
        width: "94px", minWidth: "94px", maxWidth: "94px",
        height: "26px", minHeight: "26px", maxHeight: "26px",
        borderRadius: "16px",
        border: "1px solid rgba(200, 210, 230, 0.6)",
        backdropFilter: "blur(12px)",
        WebkitBackdropFilter: "blur(12px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "flex-start",
        gap: "5px",
        padding: "0 8px",
        fontSize: "10px",
        fontWeight: 800,
        lineHeight: 1,
        textTransform: "uppercase",
        cursor: "pointer",
        flexShrink: 0,
        margin: 0,
        outline: "none",
        /* White Gel 3D */
        background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
        color: "#172B4D",
        boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
    },
    alert: {
        padding: "8px 12px",
        background: "#FFEBE6",
        color: "#BF2600",
        borderRadius: "3px",
        fontSize: "12px",
    },
    section: {
        border: "1px solid #DFE1E6",
        borderRadius: "3px",
        overflow: "hidden",
        background: "white",
    },
    sectionHeader: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
        padding: "8px 12px",
        background: "#F4F5F7",
        cursor: "pointer",
        borderBottom: "1px solid #DFE1E6",
    },
    sectionTitle: {
        fontSize: "11px",
        fontWeight: 700,
        color: "#42526E",
        margin: 0,
        textTransform: "uppercase",
    },
    refreshBtn: {
        background: "none",
        border: "none",
        color: "#6B778C",
        cursor: "pointer",
        padding: "4px",
    },
    accordionContent: {
        padding: "12px",
    },
    emptyState: {
        padding: "24px",
        textAlign: "center",
        color: "#6B778C",
        fontSize: "12px",
    },
    cardList: {
        display: "flex",
        flexDirection: "column",
        gap: "8px",
    },
    card: {
        padding: "12px",
        border: "1px solid #DFE1E6",
        borderRadius: "3px",
        background: "white",
        boxShadow: "0 1px 1px rgba(9, 30, 66, 0.25)",
    },
    cardHeader: {
        display: "flex",
        alignItems: "center",
        gap: "8px",
        marginBottom: "8px",
    },
    modelTag: {
        fontSize: "10px",
        fontWeight: 700,
        padding: "2px 8px",
        borderRadius: "16px",
        textTransform: "uppercase",
    },
    recordId: {
        fontSize: "11px",
        color: "#6B778C",
        fontWeight: 500,
    },
    quickActionBtn: {
        padding: "4px",
        background: "none",
        border: "none",
        color: "#6B778C",
        cursor: "pointer",
        borderRadius: "3px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
    },
    cardTitle: {
        fontSize: "13px",
        fontWeight: 600,
        color: "#172B4D",
        lineHeight: "1.4",
        whiteSpace: "nowrap",
        overflow: "hidden",
        textOverflow: "ellipsis",
    },
    cardFooter: {
        display: "flex",
        justifyContent: "flex-end",
        marginTop: "8px",
    },
    linkBtn: {
        fontSize: "11px",
        color: "#0052CC",
        textDecoration: "none",
        fontWeight: 600,
        display: "flex",
        alignItems: "center",
        gap: "4px",
    },
    voiceBar: {
        padding: "12px",
        background: "white",
        borderTop: "1px solid #DFE1E6",
    },
    voiceInputWrapper: {
        display: "flex",
        alignItems: "center",
        gap: "8px",
        background: "#FFFFFF",
        border: "2px solid #DFE1E6",
        borderRadius: "3px",
        padding: "6px 10px",
    },
    voiceInput: {
        flex: 1,
        background: "none",
        border: "none",
        fontSize: "13px",
        outline: "none",
        color: "#172B4D",
    },
    voiceSendBtn: {
        background: "#0052CC",
        border: "none",
        color: "white",
        borderRadius: "3px",
        width: "24px",
        height: "24px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        padding: 0,
    },
    footer: {
        padding: "8px",
        fontSize: "11px",
        textAlign: "center",
        color: "#6B778C",
        borderTop: "1px solid #DFE1E6",
        background: "#F4F5F7",
    },
};
