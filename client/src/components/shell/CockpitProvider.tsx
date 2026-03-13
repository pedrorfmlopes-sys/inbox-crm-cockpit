import React, { createContext, useContext, useEffect, useState, useRef } from "react";
import { getSelectedMessageContext, subscribeToItemChanges, getCurrentItemToken, getEmailBodyHtml, getEmailBodyText, syncManualGroupCategories, syncOdooLinkedCategory, syncOdooLinkedNotification, type OutlookAttachment, type OutlookMessageContext } from "@/office";
import { getLinks, getOdooMeta, getRelatedEmailContext, login as apiLogin, checkAuth as apiCheckAuth, registerRelevantEmail, setApiSessionToken, type LinkEntry, type OdooMeta } from "@/api";
import { getCachedSettingsSnapshot, getSettings, saveSettings, SETTINGS_UPDATED_EVENT, type CockpitSettingsV1 } from "@/settings";
import { clientLog } from "@/logger";
import { type AiTone, type AiLocale } from "@/ai/aiClient";

export type CockpitTab = "ai" | "crm" | "crm2" | "related" | "groups" | "files" | "settings";
export type StartupCheckStatus = "pending" | "running" | "success" | "warning" | "error";

type StartupCheckId = "settings" | "session" | "email" | "links" | "services";

export interface StartupCheck {
    id: StartupCheckId;
    label: string;
    detail: string;
    status: StartupCheckStatus;
}

export interface StartupNotice {
    tone: "info" | "error";
    title: string;
    details: string[];
}

export interface ConnectivityCheckResult {
    odooOk: boolean;
    openaiOk: boolean;
    geminiOk: boolean;
    summary: string;
    failures: string[];
}

interface GeminiStatusDetails {
    requested?: string;
    sanitized?: string;
    effective?: string;
    provider?: string;
}

interface GranularStatusDetails {
    openai: string | null;
    gemini: string | null;
    geminiDetails: GeminiStatusDetails | null;
}

export interface AiState {
    prompt: string;
    output: string;
    tone: AiTone;
    locale: AiLocale;
    history: Array<{ role: "user" | "assistant"; content: string }>;
    smartReplies: string[];
    action?: string;
    suggestedTo?: string[];
    suggestedCc?: string[];
    suggestedSubject?: string;
}

export interface CockpitContextType {
    tab: CockpitTab;
    setTab: (tab: CockpitTab) => void;
    ctx: OutlookMessageContext;
    bodyText: string;
    bodyHtml: string;
    meta: OdooMeta | null;
    links: LinkEntry[];
    attachments: OutlookAttachment[];
    msg: string | null;
    setMsg: (msg: string | null) => void;
    refreshLinks: () => Promise<void>;
    isLoading: boolean;
    aiState: AiState;
    setAiState: (update: Partial<AiState>) => void;
    files: Array<{ name: string; type: string; content: string }>;
    addFile: (file: { name: string; type: string; content: string }) => void;
    removeFile: (name: string) => void;
    clearFiles: () => void;
    isAuthenticated: boolean;
    connectionStatus: "none" | "success" | "error";
    granularStatus: { odoo: boolean | null; openai: boolean | null; gemini: boolean | null };
    granularStatusDetails: GranularStatusDetails;
    granularStatusString: string;
    checkConnectivity: (customModels?: any) => Promise<ConnectivityCheckResult>;
    login: (credentials: any) => Promise<void>;
    logout: () => void;
    settings: CockpitSettingsV1 | null;
    activeGroupSelection: { emailKey: string; groupId: string | null };
    setActiveGroupForCurrentEmail: (groupId: string | null) => void;
    startupChecks: StartupCheck[];
    startupNotice: StartupNotice | null;
    dismissStartupNotice: () => void;
}

// Export the context so it can be checked or used elsewhere if needed (rare)
// Using a global singleton pattern to prevent duplication in Outlook/Vite HMR
type CockpitContextSingletonHost = typeof globalThis & {
    __ICCC_COCKPIT_CONTEXT_v1__?: React.Context<CockpitContextType | undefined>;
};

const G = globalThis as CockpitContextSingletonHost;
const GK = "__ICCC_COCKPIT_CONTEXT_v1__";
const ACTIVE_TAB_STORAGE_KEY = "iccc_active_tab_v1";
const WARM_BOOT_STORAGE_KEY = "iccc_warm_boot_v1";
const WARM_BOOT_MAX_AGE_MS = 10 * 60 * 1000;
const LINKS_CACHE_PREFIX = "iccc_links_cache_v1:";
const LINKS_CACHE_MESSAGE_PREFIX = "iccc_links_cache_msg_v1:";
const LINKS_CACHE_ITEM_PREFIX = "iccc_links_cache_item_v1:";
const STARTUP_CHECK_BLUEPRINT: Array<{ id: StartupCheckId; label: string; detail: string }> = [
    { id: "settings", label: "Definições", detail: "A carregar preferências e sessão guardada..." },
    { id: "session", label: "Odoo", detail: "A validar ligação e sessão Odoo..." },
    { id: "email", label: "Email atual", detail: "A ler contexto, corpo e anexos..." },
    { id: "links", label: "Ligações", detail: "A sincronizar ligações e histórico relevante..." },
    { id: "services", label: "Serviços", detail: "A testar Odoo e motores de IA..." },
];

function createStartupChecks(): StartupCheck[] {
    return STARTUP_CHECK_BLUEPRINT.map((check) => ({ ...check, status: "pending" as StartupCheckStatus }));
}

function buildContextEmailKey(ctx: OutlookMessageContext): string {
    return [
        String(ctx.itemId || "").trim(),
        String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, ""),
        String(ctx.conversationId || "").trim(),
    ].join("|");
}

if (!G[GK]) {
    G[GK] = createContext<CockpitContextType | undefined>(undefined);
}
export const CockpitContext = G[GK] as React.Context<CockpitContextType | undefined>;

export const CockpitProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    function readPersistedTab(): CockpitTab {
        try {
            const raw = sessionStorage.getItem(ACTIVE_TAB_STORAGE_KEY);
            return raw === "ai" || raw === "crm" || raw === "crm2" || raw === "related" || raw === "groups" || raw === "files" || raw === "settings"
                ? raw
                : "ai";
        } catch {
            return "ai";
        }
    }

    function hasWarmBootHint(): boolean {
        try {
            const raw = Number(sessionStorage.getItem(WARM_BOOT_STORAGE_KEY) || "0");
            return Number.isFinite(raw) && raw > 0 && (Date.now() - raw) < WARM_BOOT_MAX_AGE_MS;
        } catch {
            return false;
        }
    }

    const warmStartupRef = useRef<boolean>(hasWarmBootHint());
    const [tab, setTab] = useState<CockpitTab>(() => readPersistedTab());
    const [ctx, setCtx] = useState<OutlookMessageContext>({});
    const [bodyText, setBodyText] = useState<string>("");
    const [bodyHtml, setBodyHtml] = useState<string>("");
    const [meta, setMeta] = useState<OdooMeta | null>(null);
    const [links, setLinks] = useState<LinkEntry[]>([]);
    const [attachments, setAttachments] = useState<OutlookAttachment[]>([]);
    const [msg, setMsg] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(() => !warmStartupRef.current);
    const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => Boolean(getCachedSettingsSnapshot()?.odooSessionToken));
    const [connectionStatus, setConnectionStatus] = useState<"none" | "success" | "error">("none");
    const [settings, setSettings] = useState<CockpitSettingsV1 | null>(() => getCachedSettingsSnapshot());
    const [activeGroupSelection, setActiveGroupSelection] = useState<{ emailKey: string; groupId: string | null }>({
        emailKey: "",
        groupId: null,
    });
    const [currentCustomGroupNames, setCurrentCustomGroupNames] = useState<string[]>([]);
    const [startupChecks, setStartupChecks] = useState<StartupCheck[]>(() => createStartupChecks());
    const [startupNoticeState, setStartupNoticeState] = useState<StartupNotice | null>(null);
    const [startupNoticeDismissed, setStartupNoticeDismissed] = useState(false);

    function resetStartupPreflight() {
        setStartupChecks(createStartupChecks());
        setStartupNoticeState(null);
        setStartupNoticeDismissed(false);
    }

    function updateStartupCheck(id: StartupCheckId, patch: Partial<StartupCheck>) {
        setStartupChecks((prev) => prev.map((check) => (check.id === id ? { ...check, ...patch, id } : check)));
    }

    function setStartupNotice(next: StartupNotice | null) {
        setStartupNoticeState(next);
        setStartupNoticeDismissed(false);
    }

    function setActiveGroupForCurrentEmail(groupId: string | null) {
        setActiveGroupSelection({
            emailKey: buildContextEmailKey(ctx),
            groupId: groupId ? String(groupId).trim() : null,
        });
    }

    function dedupeLinks(entries: LinkEntry[]): LinkEntry[] {
        const seen = new Set<string>();
        return (entries || []).filter((entry) => {
            const key = `${entry.model}:${entry.recordId ?? entry.resId}:${entry.recordName ?? entry.name ?? ""}`;
            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        });
    }

    function readCachedLinks(conversationId?: string | null, internetMessageId?: string | null, itemId?: string | null): LinkEntry[] {
        if (!conversationId && !internetMessageId && !itemId) return [];
        try {
            const sources = [
                conversationId ? localStorage.getItem(`${LINKS_CACHE_PREFIX}${conversationId}`) : null,
                internetMessageId ? localStorage.getItem(`${LINKS_CACHE_MESSAGE_PREFIX}${internetMessageId}`) : null,
                itemId ? localStorage.getItem(`${LINKS_CACHE_ITEM_PREFIX}${itemId}`) : null,
            ];
            const parsed = sources.flatMap((raw) => {
                if (!raw) return [];
                try {
                    const value = JSON.parse(raw);
                    return Array.isArray(value) ? value : [];
                } catch {
                    return [];
                }
            });
            return dedupeLinks(parsed);
        } catch {
            return [];
        }
    }

    function writeCachedLinks(conversationId: string | undefined, internetMessageId: string | undefined, itemId: string | undefined, nextLinks: LinkEntry[]) {
        try {
            if (conversationId) {
                localStorage.setItem(`${LINKS_CACHE_PREFIX}${conversationId}`, JSON.stringify(nextLinks || []));
            }
            if (internetMessageId) {
                localStorage.setItem(`${LINKS_CACHE_MESSAGE_PREFIX}${internetMessageId}`, JSON.stringify(nextLinks || []));
            }
            if (itemId) {
                localStorage.setItem(`${LINKS_CACHE_ITEM_PREFIX}${itemId}`, JSON.stringify(nextLinks || []));
            }
        } catch {
            // ignore cache failures
        }
    }

    useEffect(() => {
        const handleSettingsUpdated = (event: Event) => {
            const next = (event as CustomEvent<CockpitSettingsV1>).detail;
            if (!next) return;
            setSettings(next);
            setApiSessionToken(next.odooSessionToken || null);
        };

        window.addEventListener(SETTINGS_UPDATED_EVENT, handleSettingsUpdated as EventListener);
        return () => window.removeEventListener(SETTINGS_UPDATED_EVENT, handleSettingsUpdated as EventListener);
    }, []);

    useEffect(() => {
        setApiSessionToken(settings?.odooSessionToken || null);
    }, [settings?.odooSessionToken]);

    useEffect(() => {
        try {
            sessionStorage.setItem(ACTIVE_TAB_STORAGE_KEY, tab);
        } catch {
            // ignore storage failures
        }
    }, [tab]);

    // AI History Persistence
    const [aiCache, setAiCache] = useState<Record<string, AiState>>(() => {
        try {
            const saved = localStorage.getItem("iccc_ai_cache_v1");
            return saved ? JSON.parse(saved) : {};
        } catch { return {}; }
    });

    // Save AI cache whenever it changes
    useEffect(() => {
        try {
            localStorage.setItem("iccc_ai_cache_v1", JSON.stringify(aiCache));
        } catch (e) {
            clientLog("error", "[Cockpit] Failed to save AI cache to localStorage", e);
        }
    }, [aiCache]);

    const [currentAiState, setCurrentAiState] = useState<AiState>({
        prompt: "",
        output: "",
        tone: "neutro",
        locale: "auto",
        history: [],
        smartReplies: [],
        action: "reply",
        suggestedTo: [],
        suggestedCc: [],
        suggestedSubject: "",
    });

    const ctxLoadSeqRef = useRef(0);
    const lastItemTokenRef = useRef<string>("");
    const isLoadingInProgressRef = useRef(false);

    async function fetchPersistedLinks(messageCtx: OutlookMessageContext): Promise<{
        links: LinkEntry[];
        resolvedCtx: OutlookMessageContext;
    }> {
        const initialConversationId = String(messageCtx.conversationId || "").trim();
        const initialInternetMessageId = String(messageCtx.internetMessageId || "").trim();
        if (!initialConversationId && !initialInternetMessageId) {
            return { links: [], resolvedCtx: messageCtx };
        }

        const lookupCandidates: OutlookMessageContext[] = [
            messageCtx,
            { ...messageCtx, internetMessageId: "" },
            { ...messageCtx, conversationId: "" },
        ].filter((candidate, index, arr) => {
            const key = `${candidate.conversationId || ""}||${candidate.internetMessageId || ""}`;
            return key !== "||" && arr.findIndex((entry) => `${entry.conversationId || ""}||${entry.internetMessageId || ""}` === key) === index;
        });

        for (const candidate of lookupCandidates) {
            const candidateLinks = await getLinks(candidate.conversationId, candidate.internetMessageId, candidate.itemId).catch(() => []);
            if (candidateLinks.length) {
                return { links: candidateLinks, resolvedCtx: messageCtx };
            }
        }

        const refreshedCtx = await getSelectedMessageContext().catch(() => messageCtx);
        const refreshedConversationId = String(refreshedCtx?.conversationId || initialConversationId).trim();
        const refreshedInternetMessageId = String(refreshedCtx?.internetMessageId || initialInternetMessageId).trim();
        const refreshedCandidates: OutlookMessageContext[] = [
            {
                ...messageCtx,
                ...refreshedCtx,
                conversationId: refreshedConversationId,
                internetMessageId: refreshedInternetMessageId,
            },
            {
                ...messageCtx,
                ...refreshedCtx,
                conversationId: refreshedConversationId,
                internetMessageId: "",
            },
            {
                ...messageCtx,
                ...refreshedCtx,
                conversationId: "",
                internetMessageId: refreshedInternetMessageId,
            },
        ].filter((candidate, index, arr) => {
            const key = `${candidate.conversationId || ""}||${candidate.internetMessageId || ""}`;
            return key !== "||" && arr.findIndex((entry) => `${entry.conversationId || ""}||${entry.internetMessageId || ""}` === key) === index;
        });

        for (const candidate of refreshedCandidates) {
            const candidateLinks = await getLinks(candidate.conversationId, candidate.internetMessageId, candidate.itemId).catch(() => []);
            if (candidateLinks.length) {
                return {
                    links: candidateLinks,
                    resolvedCtx: candidate,
                };
            }
        }

        return {
            links: [],
            resolvedCtx: {
                ...messageCtx,
                ...refreshedCtx,
                conversationId: refreshedConversationId,
                internetMessageId: refreshedInternetMessageId,
            },
        };
    }

    async function loadContextAndLinks(reason?: string) {
        if (isLoadingInProgressRef.current) return;
        isLoadingInProgressRef.current = true;
        const reqId = ++ctxLoadSeqRef.current;
        const startupIssues: Array<{ detail: string; severity: "warning" | "error" }> = [];
        let startupNoticeHandled = false;
        const noteStartupIssue = (detail: string, severity: "warning" | "error" = "warning") => {
            startupIssues.push({ detail, severity });
        };
        try {
            // Only show the full-screen loading spinner on the initial load.
            // Background refreshes (poll/item-changed) update silently.
            if (reason === "init") {
                resetStartupPreflight();
                if (!warmStartupRef.current) {
                    setIsLoading(true);
                }
                updateStartupCheck("settings", { status: "running" });
            }

            // 1. Check Auth first
            const s = await getSettings();
            setSettings(s);
            if (reason === "init") {
                updateStartupCheck("settings", {
                    status: "success",
                    detail: s?.odooUrl || s?.odooDb || s?.odooLogin
                        ? "Definições guardadas carregadas."
                        : "Sem definições guardadas. O arranque continua normalmente.",
                });
                updateStartupCheck("session", {
                    status: "running",
                    detail: s.odooSessionToken
                        ? "A validar sessão Odoo guardada..."
                        : "Sem sessão ativa. Vamos continuar e pedir login se for preciso.",
                });
            }
            let currentToken = s.odooSessionToken;

            if (currentToken) {
                setApiSessionToken(currentToken);
                const authCheck = await apiCheckAuth();
                if (authCheck.ok) {
                    setIsAuthenticated(true);
                    if (authCheck.meta) setMeta(authCheck.meta);
                    if (reason === "init") {
                        updateStartupCheck("session", {
                            status: "success",
                            detail: "Sessão Odoo restaurada com sucesso.",
                        });
                    }
                } else {
                    // Session expired on server? Try auto-login if we have saved password
                    if (s.odooUrl && s.odooLogin && s.odooPassword) {
                        try {
                            const resp = await apiLogin({
                                url: s.odooUrl,
                                db: s.odooDb,
                                login: s.odooLogin,
                                password: s.odooPassword
                            });
                            if (resp.ok) {
                                currentToken = resp.token;
                                setApiSessionToken(currentToken);
                                setIsAuthenticated(true);
                                setMeta(resp.meta);
                                await saveSettings({ odooSessionToken: currentToken });
                                if (reason === "init") {
                                    updateStartupCheck("session", {
                                        status: "success",
                                        detail: "Sessão Odoo reaberta automaticamente.",
                                    });
                                }
                            } else {
                                setIsAuthenticated(false);
                                setApiSessionToken(null);
                                if (reason === "init") {
                                    noteStartupIssue("A sessão Odoo guardada expirou. A app abriu, mas vai pedir novo login.");
                                    updateStartupCheck("session", {
                                        status: "warning",
                                        detail: "Sessão expirada. Login manual necessário.",
                                    });
                                }
                            }
                        } catch {
                            setIsAuthenticated(false);
                            setApiSessionToken(null);
                            if (reason === "init") {
                                noteStartupIssue("Não foi possível restaurar a sessão Odoo automaticamente.");
                                updateStartupCheck("session", {
                                    status: "warning",
                                    detail: "Auto-login falhou. A app continua com acesso limitado.",
                                });
                            }
                        }
                    } else {
                        setIsAuthenticated(false);
                        setApiSessionToken(null);
                        if (reason === "init") {
                            noteStartupIssue("A sessão Odoo guardada já não é válida e não existem credenciais para auto-login.");
                            updateStartupCheck("session", {
                                status: "warning",
                                detail: "Sessão inválida. Login manual necessário.",
                            });
                        }
                    }
                }
            } else if (reason === "init") {
                updateStartupCheck("session", {
                    status: "success",
                    detail: "Sem sessão ativa. Podes autenticar-te manualmente quando quiseres.",
                });
            }

            if (reason === "init") {
                updateStartupCheck("email", {
                    status: "running",
                    detail: "A ler o email atual e os anexos disponíveis...",
                });
            }
            const [c, b, bh, atts] = await Promise.all([
                getSelectedMessageContext(),
                getEmailBodyText(),
                getEmailBodyHtml(),
                import("@/office").then(m => m.getAttachments())
            ]);

            if (reqId !== ctxLoadSeqRef.current) return;

            // Update token tracker to avoid redundant polls
            const freshTok = await getCurrentItemToken();
            lastItemTokenRef.current = freshTok;
            setCtx(c);
            setBodyText(b);
            setBodyHtml(bh);
            setAttachments(atts || []);
            if (reason === "init") {
                const hasEmailContext = Boolean(c.itemId || c.internetMessageId || c.conversationId || c.subject);
                if (hasEmailContext) {
                    updateStartupCheck("email", {
                        status: "success",
                        detail: c.subject
                            ? `Email pronto: ${c.subject}`
                            : "Contexto do email atual carregado.",
                    });
                } else {
                    noteStartupIssue("Não foi possível identificar um email aberto. Algumas áreas podem abrir vazias.");
                    updateStartupCheck("email", {
                        status: "warning",
                        detail: "Sem email aberto ou contexto Outlook incompleto.",
                    });
                }
                updateStartupCheck("links", {
                    status: "running",
                    detail: "A sincronizar ligações persistidas e cache local...",
                });
            }
            void registerRelevantEmail({
                itemId: c.itemId || "",
                internetMessageId: c.internetMessageId || "",
                conversationId: c.conversationId || "",
                subject: c.subject || "",
                fromEmail: c.fromEmail || "",
                fromName: c.fromName || "",
                receivedAtIso: c.receivedDateTimeIso || "",
                messageDateIso: c.receivedDateTimeIso || "",
                bodyText: b,
                bodyHtml: bh,
                attachments: (atts || []).map((attachment) => ({
                    name: attachment.name,
                    contentType: attachment.contentType,
                    size: attachment.size,
                    id: attachment.id,
                    isInline: attachment.isInline,
                    contentId: attachment.contentId,
                    content:
                        attachment.isInline || String(attachment.contentType || "").trim().toLowerCase().startsWith("image/")
                            ? String(attachment.content || "").trim()
                            : "",
                })),
            }).catch(() => {
                // best-effort central registry only
            });

            if (c.conversationId) {
                setAiCache(prev => {
                    const cached = prev[c.conversationId!];
                    if (cached) {
                        setCurrentAiState(cached);
                    } else {
                        setCurrentAiState({
                            prompt: "", output: "", tone: "neutro", locale: "auto", history: [], smartReplies: [], action: "reply", suggestedTo: [], suggestedCc: [], suggestedSubject: "",
                        });
                    }
                    return prev;
                });
            }

            clientLog("info", `[Cockpit] ctx updated (${reason || 'unknown'}) conversationId=${c.conversationId || ''}`);

            if (!c.conversationId && !c.internetMessageId) {
                const cachedByItem = readCachedLinks(undefined, undefined, c.itemId);
                setLinks(cachedByItem);
                if (reason === "init") {
                    updateStartupCheck("links", {
                        status: "success",
                        detail: cachedByItem.length
                            ? `${cachedByItem.length} ligação(ões) locais prontas para o email atual.`
                            : "Sem ligações persistidas para este email.",
                    });
                }
            } else {
                setMsg(null);
                const cachedLinks = readCachedLinks(c.conversationId, c.internetMessageId, c.itemId);
                setLinks(cachedLinks);

                try {
                    const { links: l, resolvedCtx } = await fetchPersistedLinks(c);
                    if (reqId !== ctxLoadSeqRef.current) return;
                    if (
                        resolvedCtx.conversationId !== c.conversationId ||
                        resolvedCtx.internetMessageId !== c.internetMessageId
                    ) {
                        setCtx(resolvedCtx);
                    }
                    const nextLinks = (l && l.length) ? l : cachedLinks;
                    setLinks(nextLinks || []);
                    writeCachedLinks(
                        resolvedCtx.conversationId || c.conversationId,
                        resolvedCtx.internetMessageId || c.internetMessageId,
                        resolvedCtx.itemId || c.itemId,
                        nextLinks || []
                    );
                    if (reason === "init") {
                        updateStartupCheck("links", {
                            status: "success",
                            detail: nextLinks.length
                                ? `${nextLinks.length} ligação(ões) sincronizadas com o servidor.`
                                : "Sem ligações persistidas para a conversa atual.",
                        });
                    }
                } catch (e) {
                    clientLog("error", "[Cockpit] Unexpected link load error", e);
                    if (reason === "init") {
                        noteStartupIssue("As ligações persistidas não responderam a tempo. A app abriu com o que estava em cache.");
                        updateStartupCheck("links", {
                            status: "warning",
                            detail: "Falha a sincronizar ligações. Continuamos com cache local.",
                        });
                    }
                }
            }

            // 2. Load Odoo data ONLY if authenticated
            const isDev = (window as any).location.hostname === "localhost" || (window as any).location.hostname === "127.0.0.1";
            if (s.odooSessionToken || (isDev && !isAuthenticated)) {
                try {
                    const m = !meta ? await getOdooMeta().catch(() => null) : meta;

                    if (reqId !== ctxLoadSeqRef.current) return;
                    if (m) {
                        setMeta(m);
                        setIsAuthenticated(true);
                    }
                } catch (e) {
                    clientLog("error", "[Cockpit] Unexpected Odoo load error", e);
                }
            }

            if (reason === "init") {
                updateStartupCheck("services", {
                    status: "running",
                    detail: "A validar Odoo e os motores de IA...",
                });
                const connectivity = await checkConnectivity();
                const failures = connectivity.failures;
                if (failures.length) {
                    failures.forEach((failure) => noteStartupIssue(failure, "warning"));
                    updateStartupCheck("services", {
                        status: failures.length === 3 ? "error" : "warning",
                        detail: connectivity.summary,
                    });
                } else {
                    updateStartupCheck("services", {
                        status: "success",
                        detail: connectivity.summary,
                    });
                }
            }
        } catch (e: any) {
            if (reqId !== ctxLoadSeqRef.current) return;
            clientLog("error", "[Cockpit] Fatal initialization error", e);
            if (reason === "init") {
                noteStartupIssue("O arranque não terminou limpo. A app vai abrir com o que conseguiu carregar.", "error");
                setStartupNotice({
                    tone: "error",
                    title: "Arranque concluído com falhas",
                    details: ["O carregamento inicial teve um erro inesperado.", "Algumas áreas podem abrir com dados parciais."],
                });
                startupNoticeHandled = true;
            }
        } finally {
            isLoadingInProgressRef.current = false;
            if (reqId === ctxLoadSeqRef.current) {
                if (reason === "init" && !startupNoticeHandled) {
                    setStartupNotice(
                        startupIssues.length
                            ? {
                                tone: startupIssues.some((issue) => issue.severity === "error") ? "error" : "info",
                                title: startupIssues.some((issue) => issue.severity === "error")
                                    ? "Arranque concluído com alertas"
                                    : "Arranque concluído com avisos",
                                details: startupIssues.map((issue) => issue.detail),
                            }
                            : null
                    );
                }
                if (reason === "init") {
                    warmStartupRef.current = false;
                }
                setIsLoading(false);
            }
        }
    }

    const [granularStatus, setGranularStatus] = useState<{ odoo: boolean | null; openai: boolean | null; gemini: boolean | null }>({
        odoo: null,
        openai: null,
        gemini: null
    });

    const [granularStatusDetails, setGranularStatusDetails] = useState<GranularStatusDetails>({
        openai: null,
        gemini: null,
        geminiDetails: null
    });

    const [granularStatusString, setGranularStatusString] = useState<string>("Odoo: -- | OpenAI: -- | Gemini: --");

    async function checkConnectivity(customModels?: any): Promise<ConnectivityCheckResult> {
        const { odooPing, aiSelftest, getOdooMeta } = await import("@/api");
        try {
            const [o, a] = await Promise.all([
                odooPing().catch(() => ({ ok: false })),
                (aiSelftest(customModels) as any).catch(() => ({ ok: false, openai: { ok: false, error: "Falha no pedido" }, gemini: { ok: false, error: "Falha no pedido" } }))
            ]);

            let finalOdooOk = o.ok;
            let currentMeta = meta;

            // If ping works but meta is missing, try a proactive refresh
            if (o.ok && !currentMeta?.baseUrl) {
                try {
                    const freshMeta = await getOdooMeta();
                    if (freshMeta?.baseUrl) {
                        setMeta(freshMeta);
                        currentMeta = freshMeta;
                    } else {
                        finalOdooOk = false; // Ping works but we can't get usable metadata
                    }
                } catch {
                    finalOdooOk = false;
                }
            } else if (o.ok && currentMeta?.baseUrl) {
                // All good
            } else {
                finalOdooOk = false;
            }

            const newStatus = {
                odoo: finalOdooOk,
                openai: Boolean(a.openai?.ok),
                gemini: Boolean(a.gemini?.ok)
            };
            setGranularStatus(newStatus);
            setGranularStatusDetails({
                openai: a.openai?.error || (a.openai?.ok ? `Model: ${a.openai.modelUsed || 'default'}` : null),
                gemini: a.gemini?.error || (a.gemini?.ok ? `Model: ${a.gemini.modelUsed || 'default'}` : null),
                geminiDetails: a.gemini?.ok ? {
                    requested: a.gemini.requestedModel,
                    sanitized: a.gemini.sanitizedModel,
                    effective: a.gemini.modelUsed,
                    provider: a.gemini.providerUsed
                } : null
            });
            const summary = `Odoo: ${finalOdooOk ? 'OK' : 'Erro'} | OpenAI: ${a.openai?.ok ? 'OK' : 'Erro'} | Gemini: ${a.gemini?.ok ? 'OK' : 'Erro'}`;
            setGranularStatusString(summary);
            setConnectionStatus(finalOdooOk && a.ok ? "success" : "error");
            const failures: string[] = [];
            if (!finalOdooOk) failures.push("A verificação ao Odoo falhou ou devolveu metadata incompleta.");
            if (!a.openai?.ok) failures.push(a.openai?.error || "A verificação OpenAI falhou.");
            if (!a.gemini?.ok) failures.push(a.gemini?.error || "A verificação Gemini falhou.");
            return {
                odooOk: finalOdooOk,
                openaiOk: Boolean(a.openai?.ok),
                geminiOk: Boolean(a.gemini?.ok),
                summary,
                failures,
            };
        } catch {
            setConnectionStatus("error");
            setGranularStatusString("Odoo: Error | AI: Error");
            return {
                odooOk: false,
                openaiOk: false,
                geminiOk: false,
                summary: "Odoo: Erro | OpenAI: Erro | Gemini: Erro",
                failures: ["Não foi possível concluir as verificações de conectividade."],
            };
        }
    }

    // Heartbeat logic
    useEffect(() => {
        const interval = setInterval(() => {
            console.log("[Cockpit] Heartbeat: Checking connectivity...");
            // Use current settings if available in state or just fallback to server config
            checkConnectivity();
        }, 5 * 60 * 1000); // 5 minutes

        return () => clearInterval(interval);
    }, []);

    useEffect(() => {
        loadContextAndLinks('init');

        let unsub: (() => void) | null = null;
        (async () => {
            try {
                unsub = await subscribeToItemChanges(() => loadContextAndLinks('item-changed'));
            } catch (e) {
                clientLog("warn", "[Cockpit] subscription failed", e);
            }
        })();

        const intervalId = window.setInterval(async () => {
            if (isLoadingInProgressRef.current) return;
            try {
                // Using getCurrentItemToken (sync properties + poke) 
                // as it's faster for high-frequency polling
                const freshTok = await getCurrentItemToken();

                if (freshTok && freshTok !== lastItemTokenRef.current) {
                    clientLog("info", "[Cockpit] Change detected via poll. Triggering double-reload.");

                    // Trigger immediate reload
                    loadContextAndLinks('poll-immediate');

                    // Late reload: handles Outlook bridge lag where data is partially stale right after switching
                    window.setTimeout(() => {
                        loadContextAndLinks('poll-late');
                    }, 600);
                }
            } catch { }
        }, 1000); // Increased frequency (1s) for more responsive UI

        return () => {
            unsub?.();
            window.clearInterval(intervalId);
        };
    }, []);

    // Rescue Timeout: if app is stuck loading for > 15s, force clear and show error
    useEffect(() => {
        if (!isLoading) return;
        const rescueId = setTimeout(() => {
            if (isLoading) {
                clientLog("error", "[Cockpit] Rescue timeout triggered! Loading took > 15s.");
                setMsg("A inicialização demorou demasiado tempo. Verifica a ligação ao servidor ou se o Office.js está a responder.");
                setIsLoading(false);
                isLoadingInProgressRef.current = false;
            }
        }, 15000);
        return () => clearTimeout(rescueId);
    }, [isLoading]);

    const refreshLinks = async () => {
        if (!ctx.conversationId && !ctx.internetMessageId && !ctx.itemId) return;
        try {
            const cachedLinks = readCachedLinks(ctx.conversationId, ctx.internetMessageId, ctx.itemId);
            const { links: l, resolvedCtx } = await fetchPersistedLinks(ctx);
            if (
                resolvedCtx.conversationId !== ctx.conversationId ||
                resolvedCtx.internetMessageId !== ctx.internetMessageId
            ) {
                setCtx(resolvedCtx);
            }
            const nextLinks = (l && l.length) ? l : cachedLinks;
            setLinks(nextLinks || []);
            writeCachedLinks(
                resolvedCtx.conversationId || ctx.conversationId,
                resolvedCtx.internetMessageId || ctx.internetMessageId,
                resolvedCtx.itemId || ctx.itemId,
                nextLinks || []
            );
        } catch (e: any) {
            setMsg(e?.message ?? String(e));
        }
    };

    useEffect(() => {
        const hasContextIdentity = Boolean(ctx.itemId || ctx.internetMessageId || ctx.conversationId);
        if (!hasContextIdentity) {
            setCurrentCustomGroupNames([]);
            return;
        }
        let cancelled = false;
        const payload = {
            itemId: String(ctx.itemId || "").trim(),
            internetMessageId: String(ctx.internetMessageId || "").trim(),
            conversationId: String(ctx.conversationId || "").trim(),
            subject: String(ctx.subject || "").trim(),
            fromEmail: String(ctx.fromEmail || "").trim(),
            fromName: String(ctx.fromName || "").trim(),
            receivedAtIso: String(ctx.receivedDateTimeIso || "").trim(),
            messageDateIso: String(ctx.receivedDateTimeIso || "").trim(),
        };
        getRelatedEmailContext(payload)
            .then((response) => {
                if (cancelled) return;
                const names = Array.isArray(response?.groups)
                    ? response.groups
                        .filter((group) => group.kind === "custom")
                        .map((group) => String(group.name || "").trim())
                        .filter(Boolean)
                    : [];
                setCurrentCustomGroupNames(names);
            })
            .catch(() => {
                if (!cancelled) setCurrentCustomGroupNames([]);
            });

        return () => {
            cancelled = true;
        };
    }, [ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject, links.length]);

    useEffect(() => {
        syncOdooLinkedCategory(links.length > 0).catch(() => {
            // best-effort host hint only
        });
        syncOdooLinkedNotification(links.length > 0, links.length).catch(() => {
            // best-effort host hint only
        });
        syncManualGroupCategories(currentCustomGroupNames).catch(() => {
            // best-effort host hint only
        });
    }, [ctx.itemId, currentCustomGroupNames, links.length]);

    const setAiState = (update: Partial<AiState>) => {
        if (!ctx.conversationId) return;
        setCurrentAiState(prev => {
            const newState = { ...prev, ...update };
            setAiCache(cache => ({
                ...cache,
                [ctx.conversationId!]: newState
            }));
            return newState;
        });
    };

    const [files, setFiles] = useState<Array<{ name: string; type: string; content: string }>>([]);

    // ... (rest of the state)

    const addFile = (file: { name: string; type: string; content: string }) => {
        setFiles(prev => [...prev, file]);
    };

    const removeFile = (name: string) => {
        setFiles(prev => prev.filter(f => f.name !== name));
    };

    const clearFiles = () => {
        setFiles([]);
    };

    // Reset files when conversation changes
    useEffect(() => {
        clearFiles();
    }, [ctx.conversationId]);

    const login = async (credentials: any) => {
        const resp = await apiLogin(credentials);
        if (resp.ok) {
            setApiSessionToken(resp.token);
            setIsAuthenticated(true);
            setMeta(resp.meta);
            await saveSettings({
                odooUrl: credentials.url,
                odooDb: credentials.db,
                odooLogin: credentials.login,
                odooPassword: credentials.password,
                odooSessionToken: resp.token
            });
            await loadContextAndLinks('login-success');
            await checkConnectivity();
        } else {
            throw new Error(resp.message);
        }
    };

    const logout = async () => {
        setApiSessionToken(null);
        setIsAuthenticated(false);
        setMeta(null);
        setLinks([]);
        await saveSettings({ odooSessionToken: "" });
    };

    return (
        <CockpitContext.Provider value={{
            tab, setTab, ctx, bodyText, bodyHtml, attachments, meta, links, msg, setMsg, refreshLinks, isLoading,
            aiState: currentAiState,
            setAiState,
            files, addFile, removeFile, clearFiles,
            isAuthenticated, connectionStatus, granularStatus, granularStatusDetails, granularStatusString, checkConnectivity, login, logout,
            settings,
            activeGroupSelection,
            setActiveGroupForCurrentEmail,
            startupChecks,
            startupNotice: startupNoticeDismissed ? null : startupNoticeState,
            dismissStartupNotice: () => setStartupNoticeDismissed(true),
        }}>
            {children}
        </CockpitContext.Provider>
    );
};

export const useCockpit = () => {
    const context = useContext(CockpitContext);
    if (context === undefined) {
        throw new Error("useCockpit must be used within a CockpitProvider");
    }
    return context;
};
