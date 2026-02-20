import React, { createContext, useContext, useEffect, useState, useRef } from "react";
import { getSelectedMessageContext, subscribeToItemChanges, getCurrentItemToken, getEmailBodyText, type OutlookMessageContext } from "@/office";
import { getLinks, getOdooMeta, type LinkEntry, type OdooMeta } from "@/api";
import { type AiTone, type AiLocale } from "@/ai/aiClient";
import { clientLog } from "@/logger";

export type CockpitTab = "ai" | "crm" | "files" | "settings";

export interface AiState {
    prompt: string;
    output: string;
    tone: AiTone;
    locale: AiLocale;
    history: Array<{ role: "user" | "assistant"; content: string }>;
}

export interface CockpitContextType {
    tab: CockpitTab;
    setTab: (tab: CockpitTab) => void;
    ctx: OutlookMessageContext;
    bodyText: string;
    meta: OdooMeta | null;
    links: LinkEntry[];
    msg: string | null;
    setMsg: (msg: string | null) => void;
    refreshLinks: () => Promise<void>;
    isLoading: boolean;
    aiState: AiState;
    setAiState: (update: Partial<AiState>) => void;
    files: Array<{ name: string; type: string; content: string }>;
    addFile: (file: { name: string; type: string; content: string }) => void;
    removeFile: (name: string) => void;
}

// Export the context so it can be checked or used elsewhere if needed (rare)
// Using a global singleton pattern to prevent duplication in Outlook/Vite HMR
const G = (typeof window !== 'undefined' ? window : {}) as any;
const GK = "__ICCC_COCKPIT_CONTEXT_v1__";

if (!G[GK]) {
    G[GK] = createContext<CockpitContextType | undefined>(undefined);
}
export const CockpitContext = G[GK];

export const CockpitProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const [tab, setTab] = useState<CockpitTab>("ai");
    const [ctx, setCtx] = useState<OutlookMessageContext>({});
    const [bodyText, setBodyText] = useState<string>("");
    const [meta, setMeta] = useState<OdooMeta | null>(null);
    const [links, setLinks] = useState<LinkEntry[]>([]);
    const [msg, setMsg] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState(true);

    const [aiCache, setAiCache] = useState<Record<string, AiState>>({});
    const [currentAiState, setCurrentAiState] = useState<AiState>({
        prompt: "",
        output: "",
        tone: "neutro",
        locale: "auto",
        history: [],
    });

    const ctxLoadSeqRef = useRef(0);
    const lastItemTokenRef = useRef<string>("");
    const isLoadingInProgressRef = useRef(false);

    async function loadContextAndLinks(reason?: string) {
        if (isLoadingInProgressRef.current) return;
        isLoadingInProgressRef.current = true;
        const reqId = ++ctxLoadSeqRef.current;
        try {
            setIsLoading(true);
            const [c, b] = await Promise.all([
                getSelectedMessageContext(),
                getEmailBodyText(),
            ]);

            if (reqId !== ctxLoadSeqRef.current) return;
            setCtx(c);
            setBodyText(b);

            if (c.conversationId) {
                setAiCache(prev => {
                    const cached = prev[c.conversationId!];
                    if (cached) {
                        setCurrentAiState(cached);
                    } else {
                        // Reset to clean state if no cache
                        setCurrentAiState({
                            prompt: "",
                            output: "",
                            tone: "neutro",
                            locale: "auto",
                            history: [],
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

            // Load Odoo data optionally
            try {
                const [l, m] = await Promise.all([
                    getLinks(c.conversationId).catch(err => {
                        clientLog("warn", "[Cockpit] getLinks failed:", err);
                        return [];
                    }),
                    !meta ? getOdooMeta().catch(err => {
                        clientLog("warn", "[Cockpit] getOdooMeta failed:", err);
                        return null;
                    }) : Promise.resolve(meta)
                ]);

                if (reqId !== ctxLoadSeqRef.current) return;
                setLinks(l || []);
                if (m) setMeta(m);
            } catch (e) {
                clientLog("error", "[Cockpit] Unexpected load error", e);
            }
        } catch (e: any) {
            if (reqId !== ctxLoadSeqRef.current) return;
            // Only fatal errors (like Office.js fails) should setMsg
            clientLog("error", "[Cockpit] Fatal initialization error", e);
        } finally {
            isLoadingInProgressRef.current = false;
            if (reqId === ctxLoadSeqRef.current) setIsLoading(false);
        }
    }

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
                const tok = await getCurrentItemToken();
                if (tok && tok !== lastItemTokenRef.current) {
                    lastItemTokenRef.current = tok;
                    loadContextAndLinks('poll');
                }
            } catch { }
        }, 3000);

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

    // Reset files when conversation changes
    useEffect(() => {
        setFiles([]);
    }, [ctx.conversationId]);

    return (
        <CockpitContext.Provider value={{
            tab, setTab, ctx, bodyText, meta, links, msg, setMsg, refreshLinks, isLoading,
            aiState: currentAiState,
            setAiState,
            files, addFile, removeFile
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
