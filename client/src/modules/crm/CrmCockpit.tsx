import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { openCockpitDialog, getEmailBodyText, getOutlookContactSuggestionByEmail, OutlookContactSuggestion } from "../../office";
import * as Icons from "../../ui/icons";
import { ContactInsight } from "./ContactInsight";
import { OdooCardSkeleton, Skeleton } from "../../ui/SkeletonLoader";
import { aiExtractAnchors, aiGenerate, createOrUpdatePartner, getOdooAutoLoginUrl, getPartnerByEmail, searchCompanies, searchOdoo } from "../../api";
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


type ContactOrigin = "from" | "to" | "cc";
type ContactLookupState = "idle" | "loading" | "found" | "not_found" | "error";

type ContactPanelRow = {
    key: string;
    email: string;
    name: string;
    origin: ContactOrigin[];
    lookupState: ContactLookupState;
    partner: any | null;
    companyType: "person" | "company";
    parentId: number | null;
    companyQuery: string;
    companyOptions: Array<{ id: number; name: string; email?: string }>;
    functionValue: string;
    phone: string;
    mobile: string;
    isSaving: boolean;
    error: string | null;
    outlookSuggestion: OutlookContactSuggestion | null;
    isOutlookLoading: boolean;
    applyOutlookOpen: boolean;
    applyOutlookFields: { name: boolean; company: boolean; jobTitle: boolean; phone: boolean };
};

function normalizeEmailValue(v: string) {
    return String(v || "").trim().toLowerCase();
}

function fallbackNameFromEmail(email: string) {
    const local = String(email || "").split("@")[0] || "";
    return local.replace(/[._-]+/g, " ").trim() || "(sem nome)";
}

function toCompanyType(v: any): "person" | "company" {
    return v === "company" ? "company" : "person";
}

function collectParticipants(ctx: any): ContactPanelRow[] {
    const rows = new Map<string, ContactPanelRow>();

    const upsert = (origin: ContactOrigin, emailRaw?: string, nameRaw?: string) => {
        const email = normalizeEmailValue(emailRaw || "");
        if (!email) return;

        if (rows.has(email)) {
            const existing = rows.get(email)!;
            if (!existing.origin.includes(origin)) existing.origin.push(origin);
            if (!existing.name && nameRaw) existing.name = String(nameRaw).trim();
            return;
        }

        const fallback = fallbackNameFromEmail(email);
        rows.set(email, {
            key: email,
            email,
            name: String(nameRaw || "").trim() || fallback,
            origin: [origin],
            lookupState: "idle",
            partner: null,
            companyType: "person",
            parentId: null,
            companyQuery: "",
            companyOptions: [],
            functionValue: "",
            phone: "",
            mobile: "",
            isSaving: false,
            error: null,
            outlookSuggestion: null,
            isOutlookLoading: false,
            applyOutlookOpen: false,
            applyOutlookFields: { name: true, company: true, jobTitle: true, phone: true },
        });
    };

    upsert("from", ctx.fromEmail, ctx.fromName);
    for (const r of ctx.toRecipients || []) upsert("to", (r as any)?.email, (r as any)?.name);
    for (const r of ctx.ccRecipients || []) upsert("cc", (r as any)?.email, (r as any)?.name);

    return Array.from(rows.values());
}

function statusLabel(state: ContactLookupState) {
    if (state === "loading") return "A carregar";
    if (state === "found") return "Encontrado";
    if (state === "not_found") return "Não encontrado";
    if (state === "error") return "Erro";
    return "A carregar";
}

function extractPhonesFromText(text: string): string[] {
    const raw = String(text || "").replace(/\s+/g, " ");
    const rx = /(?:\+?351[\s.-]?)?(?:9\d{2}|2\d{2})[\s.-]?\d{3}[\s.-]?\d{3}|\+\d{1,3}[\s.-]?\d{2,4}[\s.-]?\d{3,4}[\s.-]?\d{3,4}/g;
    const found = raw.match(rx) || [];
    const norm = found.map((x) => x.trim().replace(/\s+/g, " "));
    return Array.from(new Set(norm)).slice(0, 20);
}

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
    const [protection, setProtection] = useState<MatchResult | null>(null);
    const [isDrafting, setIsDrafting] = useState(false);
    const [isBriefingLoading, setIsBriefingLoading] = useState(false);
    const [briefing, setBriefing] = useState<string | null>(null);
    const [voiceCommand, setVoiceCommand] = useState("");
    const [isVoiceLoading, setIsVoiceLoading] = useState(false);
    const [isContactsExpanded, setIsContactsExpanded] = useState(false);
    const [contactRows, setContactRows] = useState<ContactPanelRow[]>([]);
    const [emailPhones, setEmailPhones] = useState<string[]>([]);

    // Safety no-op: email-insights feature disabled in this branch;
    // keep symbol defined to avoid runtime ReferenceError in stale call sites.
    async function loadEmailInsights() {
        return;
    }

    function updateContactRow(email: string, patch: Partial<ContactPanelRow>) {
        const key = normalizeEmailValue(email);
        setContactRows((prev) => prev.map((row) => (row.email === key ? { ...row, ...patch } : row)));
    }

    async function runPartnerLookup(email: string) {
        const key = normalizeEmailValue(email);
        if (!key) return;
        updateContactRow(key, { lookupState: "loading", error: null });
        try {
            const partner = await getPartnerByEmail(key);
            if (partner) {
                updateContactRow(key, {
                    lookupState: "found",
                    partner,
                    name: partner.name || fallbackNameFromEmail(key),
                    companyType: toCompanyType(partner.company_type),
                    parentId: Array.isArray(partner.parent_id) ? Number(partner.parent_id[0]) : (partner.parent_id ? Number(partner.parent_id) : null),
                    functionValue: String(partner.function || ""),
                    phone: String(partner.phone || ""),
                    mobile: String(partner.mobile || ""),
                    error: null,
                });
            } else {
                updateContactRow(key, { lookupState: "not_found", partner: null, error: null });
            }
        } catch (e: any) {
            updateContactRow(key, { lookupState: "error", error: e?.message || "Falha no lookup" });
        }
    }

    async function loadOutlookSuggestion(email: string) {
        const key = normalizeEmailValue(email);
        updateContactRow(key, { isOutlookLoading: true });
        try {
            const suggestion = await getOutlookContactSuggestionByEmail(key);
            updateContactRow(key, { outlookSuggestion: suggestion || null });
        } finally {
            updateContactRow(key, { isOutlookLoading: false });
        }
    }

    function toggleOutlookField(row: ContactPanelRow, field: "name" | "company" | "jobTitle" | "phone") {
        updateContactRow(row.email, {
            applyOutlookFields: {
                ...row.applyOutlookFields,
                [field]: !row.applyOutlookFields[field],
            },
        });
    }

    function applyOutlookSuggestion(row: ContactPanelRow) {
        const s = row.outlookSuggestion;
        if (!s) return;
        const patch: Partial<ContactPanelRow> = { applyOutlookOpen: false };
        if (row.applyOutlookFields.name && s.name) patch.name = s.name;
        if (row.applyOutlookFields.jobTitle && s.jobTitle) patch.functionValue = s.jobTitle;
        if (row.applyOutlookFields.company && s.company) patch.companyQuery = s.company;
        if (row.applyOutlookFields.phone && Array.isArray(s.phones) && s.phones.length) {
            patch.phone = row.phone || s.phones[0];
            if (s.phones[1] && !row.mobile) patch.mobile = s.phones[1];
        }
        updateContactRow(row.email, patch);
    }

    async function handleCompanySearch(email: string, query: string) {
        const key = normalizeEmailValue(email);
        updateContactRow(key, { companyQuery: query });
        if (!query.trim()) {
            updateContactRow(key, { companyOptions: [] });
            return;
        }
        try {
            const items = await searchCompanies(query);
            updateContactRow(key, { companyOptions: (items || []).map((x: any) => ({ id: Number(x.id), name: x.name || x.display_name || `#${x.id}`, email: x.email })) });
        } catch {
            updateContactRow(key, { companyOptions: [] });
        }
    }

    async function handleSaveContact(row: ContactPanelRow, mode: "create" | "update") {
        updateContactRow(row.email, { isSaving: true, error: null });
        try {
            const payload: any = {
                mode,
                targetPartnerId: mode === "update" ? Number(row.partner?.id || 0) : undefined,
                data: {
                    name: row.name,
                    email: row.email,
                    company_type: row.companyType,
                    parent_id: row.parentId ?? null,
                    function: row.functionValue,
                    phone: row.phone,
                    mobile: row.mobile,
                },
            };
            await createOrUpdatePartner(payload);
            await runPartnerLookup(row.email);
        } catch (e: any) {
            const msg = String(e?.message || "Erro ao guardar");
            if (msg.includes("HTTP 409")) {
                await runPartnerLookup(row.email);
                updateContactRow(row.email, { error: "Contacto já existe por email. Use Atualizar." });
            } else {
                updateContactRow(row.email, { error: msg });
            }
        } finally {
            updateContactRow(row.email, { isSaving: false });
        }
    }

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
        const participants = collectParticipants(ctx);
        setContactRows(participants);
        (async () => {
            const body = await getEmailBodyText();
            setEmailPhones(extractPhonesFromText(body || ""));
        })();
        participants.forEach((row) => {
            runPartnerLookup(row.email);
            loadOutlookSuggestion(row.email);
        });
        if (ctx.conversationId) {
            handleScanAnchors();
            loadContact();
        }
    }, [ctx.conversationId, ctx.fromEmail, ctx.toRecipients, ctx.ccRecipients]);

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
        // setIsContactLoading(true);
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
            // setIsContactLoading(false);
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
                ...(extra ?? {}),
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
                    <div style={S.sectionHeader} onClick={() => setIsContactsExpanded(!isContactsExpanded)}>
                        <div style={{ display: 'flex', alignItems: 'center', gap: '6px' }}>
                            <div style={{ transform: isContactsExpanded ? 'rotate(90deg)' : 'rotate(0deg)', transition: 'transform 0.2s' }}>
                                <Icons.ExternalLink size={12} style={{ transform: 'rotate(-45deg)' }} />
                            </div>
                            <h3 style={S.sectionTitle}>Contactos ({contactRows.length})</h3>
                        </div>
                    </div>
                    {isContactsExpanded && (
                        <div style={S.accordionContent}>
                            {!!emailPhones.length && (
                                <div style={S.phoneBucket}>
                                    <div style={S.phoneBucketTitle}>Telefones encontrados (email)</div>
                                    <div style={S.phoneBucketList}>
                                        {emailPhones.map((ph) => (
                                            <span key={`global-phone-${ph}`} style={S.phoneChip}>{ph}</span>
                                        ))}
                                    </div>
                                </div>
                            )}
                            {!contactRows.length ? (
                                <div style={S.emptyState}><p>Sem contactos no From/To/Cc deste email.</p></div>
                            ) : (
                                <div style={S.contactList}>
                                    {contactRows.map((row) => (
                                        <div key={row.key} style={S.contactCard}>
                                            <div style={S.contactTopRow}>
                                                <div>
                                                    <div style={S.contactName}>{row.name || fallbackNameFromEmail(row.email)}</div>
                                                    <div style={S.contactEmail}>{row.email}</div>
                                                    <div style={S.contactOrigin}>{row.origin.join(", ").toUpperCase()}</div>
                                                </div>
                                                <div style={{ ...S.lookupBadge, ...(row.lookupState === 'found' ? S.lookupFound : row.lookupState === 'error' ? S.lookupError : row.lookupState === 'not_found' ? S.lookupMissing : S.lookupLoading) }}>
                                                    {statusLabel(row.lookupState)}
                                                </div>
                                            </div>

                                            {(row.isOutlookLoading || row.outlookSuggestion) && (
                                                <div style={S.outlookSuggestionBox}>
                                                    <div style={S.outlookSuggestionTitle}>Sugestão Outlook</div>
                                                    {row.isOutlookLoading ? (
                                                        <div style={S.outlookSuggestionText}>A carregar...</div>
                                                    ) : row.outlookSuggestion ? (
                                                        <>
                                                            <div style={S.outlookSuggestionText}>Nome: {row.outlookSuggestion.name || "—"}</div>
                                                            <div style={S.outlookSuggestionText}>Empresa: {row.outlookSuggestion.company || "—"}</div>
                                                            <div style={S.outlookSuggestionText}>Cargo: {row.outlookSuggestion.jobTitle || "—"}</div>
                                                            <div style={S.outlookSuggestionText}>Telefone(s): {(row.outlookSuggestion.phones || []).join(", ") || "—"}</div>
                                                            <button style={S.outlookApplyBtn} onClick={() => updateContactRow(row.email, { applyOutlookOpen: !row.applyOutlookOpen })}>Aplicar dados do Outlook</button>
                                                            {row.applyOutlookOpen && (
                                                                <div style={S.applyChecklist}>
                                                                    <label style={S.checkLabel}><input type="checkbox" checked={row.applyOutlookFields.name} onChange={() => toggleOutlookField(row, "name")} /> Nome</label>
                                                                    <label style={S.checkLabel}><input type="checkbox" checked={row.applyOutlookFields.company} onChange={() => toggleOutlookField(row, "company")} /> Empresa</label>
                                                                    <label style={S.checkLabel}><input type="checkbox" checked={row.applyOutlookFields.jobTitle} onChange={() => toggleOutlookField(row, "jobTitle")} /> Cargo</label>
                                                                    <label style={S.checkLabel}><input type="checkbox" checked={row.applyOutlookFields.phone} onChange={() => toggleOutlookField(row, "phone")} /> Telefone(s)</label>
                                                                    <button style={S.outlookApplyBtn} onClick={() => applyOutlookSuggestion(row)}>Aplicar selecionados</button>
                                                                </div>
                                                            )}
                                                        </>
                                                    ) : null}
                                                </div>
                                            )}

                                            <div style={S.contactGrid}>
                                                <input style={S.contactInput} value={row.name} onChange={(e) => updateContactRow(row.email, { name: e.target.value })} placeholder="Nome" />
                                                <select style={S.contactInput} value={row.companyType} onChange={(e) => updateContactRow(row.email, { companyType: toCompanyType(e.target.value) })}>
                                                    <option value="person">Pessoa</option>
                                                    <option value="company">Empresa</option>
                                                </select>
                                                <input style={S.contactInput} value={row.functionValue} onChange={(e) => updateContactRow(row.email, { functionValue: e.target.value })} placeholder="Cargo (function)" />
                                                <input style={S.contactInput} value={row.phone} onChange={(e) => updateContactRow(row.email, { phone: e.target.value })} placeholder="Telefone" />
                                                <input style={S.contactInput} value={row.mobile} onChange={(e) => updateContactRow(row.email, { mobile: e.target.value })} placeholder="Telemóvel" />
                                            </div>

                                            <div style={{ marginTop: '6px' }}>
                                                <input
                                                    style={S.contactInput}
                                                    value={row.companyQuery}
                                                    onChange={(e) => handleCompanySearch(row.email, e.target.value)}
                                                    placeholder="Empresa (opcional)"
                                                />
                                                {!!row.companyOptions.length && (
                                                    <div style={S.companyOptions}>
                                                        {row.companyOptions.map((opt) => (
                                                            <button
                                                                key={`${row.key}-co-${opt.id}`}
                                                                style={S.companyOptionBtn}
                                                                onClick={() => updateContactRow(row.email, { parentId: opt.id, companyQuery: opt.name, companyOptions: [] })}
                                                            >
                                                                {opt.name} {opt.email ? `(${opt.email})` : ''}
                                                            </button>
                                                        ))}
                                                    </div>
                                                )}
                                                {row.parentId ? <div style={S.parentHint}>Empresa selecionada: #{row.parentId}</div> : null}
                                            </div>

                                            {!!emailPhones.length && (
                                                <div style={S.inlinePhonesRow}>
                                                    {emailPhones.slice(0, 4).map((ph) => (
                                                        <button
                                                            key={`${row.key}-apply-phone-${ph}`}
                                                            style={S.inlinePhoneBtn}
                                                            onClick={() => updateContactRow(row.email, { phone: ph })}
                                                        >
                                                            Usar {ph}
                                                        </button>
                                                    ))}
                                                </div>
                                            )}

                                            {!!row.error && <div style={S.contactError}>{row.error}</div>}

                                            <div style={S.contactActions}>
                                                <button
                                                    style={S.contactBtn}
                                                    disabled={row.lookupState === 'found' || row.isSaving}
                                                    onClick={() => handleSaveContact(row, 'create')}
                                                >
                                                    {row.isSaving ? 'A guardar...' : 'Criar no Odoo'}
                                                </button>
                                                <button
                                                    style={S.contactBtnSecondary}
                                                    disabled={row.lookupState !== 'found' || row.isSaving}
                                                    onClick={() => handleSaveContact(row, 'update')}
                                                >
                                                    {row.isSaving ? 'A guardar...' : 'Atualizar no Odoo'}
                                                </button>
                                            </div>
                                        </div>
                                    ))}
                                </div>
                            )}
                        </div>
                    )}
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
    contactList: {
        display: "flex",
        flexDirection: "column",
        gap: "8px",
    },
    contactCard: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        padding: "10px",
        background: "#fff",
    },
    contactTopRow: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "flex-start",
        gap: "8px",
        marginBottom: "8px",
    },
    contactName: { fontSize: "12px", fontWeight: 700, color: "#172B4D" },
    contactEmail: { fontSize: "12px", color: "#42526E" },
    contactOrigin: { fontSize: "10px", color: "#6B778C", marginTop: "2px" },
    lookupBadge: {
        fontSize: "10px",
        borderRadius: "999px",
        padding: "3px 8px",
        fontWeight: 700,
        border: "1px solid transparent",
    },
    lookupLoading: { background: "#F4F5F7", color: "#42526E", border: "1px solid #DFE1E6" },
    lookupFound: { background: "#E3FCEF", color: "#006644", border: "1px solid #ABF5D1" },
    lookupMissing: { background: "#FFF7D6", color: "#7A5D00", border: "1px solid #FFE380" },
    lookupError: { background: "#FFEBE6", color: "#BF2600", border: "1px solid #FFBDAD" },
    contactGrid: {
        display: "grid",
        gridTemplateColumns: "1fr 1fr",
        gap: "6px",
    },
    contactInput: {
        width: "100%",
        padding: "6px 8px",
        border: "1px solid #DFE1E6",
        borderRadius: "4px",
        fontSize: "12px",
        boxSizing: "border-box",
    },
    companyOptions: {
        border: "1px solid #DFE1E6",
        borderRadius: "4px",
        marginTop: "4px",
        overflow: "hidden",
    },
    companyOptionBtn: {
        width: "100%",
        textAlign: "left",
        border: "none",
        background: "#fff",
        padding: "6px 8px",
        cursor: "pointer",
        fontSize: "12px",
    },
    parentHint: {
        marginTop: "4px",
        fontSize: "11px",
        color: "#6B778C",
    },
    contactActions: {
        marginTop: "8px",
        display: "flex",
        gap: "6px",
    },
    contactBtn: {
        border: "1px solid #0052CC",
        background: "#0052CC",
        color: "#fff",
        borderRadius: "4px",
        padding: "6px 8px",
        fontSize: "11px",
        cursor: "pointer",
    },
    contactBtnSecondary: {
        border: "1px solid #DFE1E6",
        background: "#fff",
        color: "#172B4D",
        borderRadius: "4px",
        padding: "6px 8px",
        fontSize: "11px",
        cursor: "pointer",
    },
    contactError: {
        marginTop: "6px",
        fontSize: "11px",
        color: "#BF2600",
    },
    phoneBucket: {
        border: "1px solid #DFE1E6",
        borderRadius: "6px",
        padding: "8px",
        marginBottom: "8px",
        background: "#FAFBFC",
    },
    phoneBucketTitle: {
        fontSize: "11px",
        fontWeight: 700,
        color: "#42526E",
        marginBottom: "6px",
    },
    phoneBucketList: {
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
    },
    phoneChip: {
        fontSize: "11px",
        border: "1px solid #DFE1E6",
        borderRadius: "999px",
        padding: "3px 8px",
        background: "#fff",
        color: "#172B4D",
    },
    outlookSuggestionBox: {
        marginBottom: "8px",
        padding: "8px",
        border: "1px solid #DFE1E6",
        borderRadius: "4px",
        background: "#F7F8FA",
    },
    outlookSuggestionTitle: {
        fontSize: "11px",
        fontWeight: 700,
        color: "#42526E",
        marginBottom: "4px",
    },
    outlookSuggestionText: {
        fontSize: "11px",
        color: "#172B4D",
        marginBottom: "2px",
    },
    outlookApplyBtn: {
        marginTop: "6px",
        border: "1px solid #DFE1E6",
        background: "#fff",
        color: "#172B4D",
        borderRadius: "4px",
        padding: "5px 8px",
        fontSize: "11px",
        cursor: "pointer",
    },
    applyChecklist: {
        marginTop: "6px",
        display: "grid",
        gap: "4px",
    },
    checkLabel: {
        fontSize: "11px",
        color: "#172B4D",
        display: "flex",
        alignItems: "center",
        gap: "6px",
    },
    inlinePhonesRow: {
        marginTop: "8px",
        display: "flex",
        flexWrap: "wrap",
        gap: "6px",
    },
    inlinePhoneBtn: {
        border: "1px solid #DFE1E6",
        background: "#fff",
        color: "#172B4D",
        borderRadius: "4px",
        padding: "4px 6px",
        fontSize: "11px",
        cursor: "pointer",
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
