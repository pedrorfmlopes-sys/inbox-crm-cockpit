import React, { createContext, useContext, useEffect, useState, useRef } from "react";
import { getSelectedMessageContext, subscribeToItemChanges, getCurrentItemToken, getEmailBodyText, type OutlookMessageContext } from "@/office";
import { getLinks, getOdooMeta, login as apiLogin, checkAuth as apiCheckAuth, setApiSessionToken, type LinkEntry, type OdooMeta } from "@/api";
import { getSettings, saveSettings, SETTINGS_UPDATED_EVENT, type CockpitSettingsV1 } from "@/settings";
import { clientLog } from "@/logger";
import { type AiTone, type AiLocale } from "@/ai/aiClient";

export type CockpitTab = "ai" | "crm" | "files" | "settings";

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
    meta: OdooMeta | null;
    links: LinkEntry[];
    attachments: Array<{ name: string; contentType: string; content: string }>;
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
    checkConnectivity: (customModels?: any) => Promise<void>;
    login: (credentials: any) => Promise<void>;
    logout: () => void;
    settings: CockpitSettingsV1 | null;
}

// Export the context so it can be checked or used elsewhere if needed (rare)
// Using a global singleton pattern to prevent duplication in Outlook/Vite HMR
type CockpitContextSingletonHost = typeof globalThis & {
    __ICCC_COCKPIT_CONTEXT_v1__?: React.Context<CockpitContextType | undefined>;
};

const G = globalThis as CockpitContextSingletonHost;
const GK = "__ICCC_COCKPIT_CONTEXT_v1__";

if (!G[GK]) {
    G[GK] = createContext<CockpitContextType | undefined>(undefined);
}
export const CockpitContext = G[GK] as React.Context<CockpitContextType | undefined>;

export const CockpitProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [tab, setTab] = useState<CockpitTab>("ai");
    const [ctx, setCtx] = useState<OutlookMessageContext>({});
    const [bodyText, setBodyText] = useState<string>("");
    const [meta, setMeta] = useState<OdooMeta | null>(null);
    const [links, setLinks] = useState<LinkEntry[]>([]);
    const [attachments, setAttachments] = useState<Array<{ name: string; contentType: string; content: string }>>([]);
    const [msg, setMsg] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState(true);
    const [isAuthenticated, setIsAuthenticated] = useState(false);
    const [connectionStatus, setConnectionStatus] = useState<"none" | "success" | "error">("none");
    const [settings, setSettings] = useState<CockpitSettingsV1 | null>(null);

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

    async function loadContextAndLinks(reason?: string) {
        if (isLoadingInProgressRef.current) return;
        isLoadingInProgressRef.current = true;
        const reqId = ++ctxLoadSeqRef.current;
        try {
            // Only show the full-screen loading spinner on the initial load.
            // Background refreshes (poll/item-changed) update silently.
            if (reason === 'init') setIsLoading(true);

            // 1. Check Auth first
            const s = await getSettings();
            setSettings(s);
            let currentToken = s.odooSessionToken;

            if (currentToken) {
                setApiSessionToken(currentToken);
                const authCheck = await apiCheckAuth();
                if (authCheck.ok) {
                    setIsAuthenticated(true);
                    if (authCheck.meta) setMeta(authCheck.meta);
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
                            } else {
                                setIsAuthenticated(false);
                                setApiSessionToken(null);
                            }
                        } catch {
                            setIsAuthenticated(false);
                            setApiSessionToken(null);
                        }
                    } else {
                        setIsAuthenticated(false);
                        setApiSessionToken(null);
                    }
                }
            }

            const [c, b, atts] = await Promise.all([
                getSelectedMessageContext(),
                getEmailBodyText(),
                import("@/office").then(m => m.getAttachments())
            ]);

            if (reqId !== ctxLoadSeqRef.current) return;

            // Update token tracker to avoid redundant polls
            const freshTok = await getCurrentItemToken();
            lastItemTokenRef.current = freshTok;
            setCtx(c);
            setBodyText(b);
            setAttachments(atts || []);

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

            if (!c.conversationId) {
                setLinks([]);
                setIsLoading(false);
                return;
            }

            setMsg(null);

            try {
                const l = await getLinks(c.conversationId).catch(() => []);
                if (reqId !== ctxLoadSeqRef.current) return;
                setLinks(l || []);
            } catch (e) {
                clientLog("error", "[Cockpit] Unexpected link load error", e);
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
        } catch (e: any) {
            if (reqId !== ctxLoadSeqRef.current) return;
            clientLog("error", "[Cockpit] Fatal initialization error", e);
        } finally {
            isLoadingInProgressRef.current = false;
            if (reqId === ctxLoadSeqRef.current) setIsLoading(false);
            // Run connectivity check silently on start
            if (reqId === ctxLoadSeqRef.current && reason === 'init') {
                checkConnectivity();
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

    async function checkConnectivity(customModels?: any) {
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
            setGranularStatusString(`Odoo: ${finalOdooOk ? 'OK' : 'Error'} | OpenAI: ${a.openai?.ok ? 'OK' : 'Error'} | Gemini: ${a.gemini?.ok ? 'OK' : 'Error'}`);
            setConnectionStatus(finalOdooOk && a.ok ? "success" : "error");
        } catch {
            setConnectionStatus("error");
            setGranularStatusString("Odoo: Error | AI: Error");
        }
    }

    // Heartbeat logic
    useEffect(() => {
        // Initial run
        checkConnectivity();

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
        if (!ctx.conversationId) return;
        try {
            const l = await getLinks(ctx.conversationId);
            setLinks(l);
        } catch (e: any) {
            setMsg(e?.message ?? String(e));
        }
    };

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
            tab, setTab, ctx, bodyText, attachments, meta, links, msg, setMsg, refreshLinks, isLoading,
            aiState: currentAiState,
            setAiState,
            files, addFile, removeFile, clearFiles,
            isAuthenticated, connectionStatus, granularStatus, granularStatusDetails, granularStatusString, checkConnectivity, login, logout,
            settings
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
