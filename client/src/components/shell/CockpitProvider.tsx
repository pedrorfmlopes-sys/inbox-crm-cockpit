import React, { createContext, useContext, useEffect, useState, useRef, useCallback } from "react";
import { executeCurrentItemOutlookCategorySync, executeOutlookCategorySourceSync, getActiveOutlookCategoryOperation, getSelectedMessageContext, subscribeToItemChanges, getCurrentItemToken, getEmailBodyHtml, getEmailBodyText, getManagedOutlookCategorySnapshot, openAppSettings, OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT, OUTLOOK_CATEGORY_SYNC_REQUEST_EVENT, OUTLOOK_CATEGORY_SYNC_REQUEST_STORAGE_KEY, readPendingOutlookCategorySyncRequest, syncOdooLinkedNotification, waitForStableSelectedMessageContext, type OutlookAttachment, type OutlookMessageContext } from "@/office";
import { getLinks, getOdooMeta, getRelatedEmailContext, login as apiLogin, checkAuth as apiCheckAuth, registerRelevantEmail, setApiSessionToken, type AuthCheckResponse, type LinkEntry, type OdooMeta } from "@/api";
import { getCachedSettingsSnapshot, getSettings, saveSettings, SETTINGS_UPDATED_EVENT, type CockpitSettingsV1 } from "@/settings";
import { clientLog } from "@/logger";
import { type AiTone, type AiLocale } from "@/ai/aiClient";
import { getGroupAttachmentStorageOptions } from "@/modules/crm/groups-v1/storage/resolveStorageMode";
import {
    areOutlookCategorySourcesEqual,
    buildOutlookCategoryPlan,
    buildOutlookCategorySourceFromRelatedContext,
    getOutlookCategoryPlanSignature,
    ODOO_LINKED_CATEGORY,
    type OutlookCategorySource,
} from "@/outlookCategories";

export type CockpitTab = "ai" | "crm" | "crm2" | "related" | "groups" | "files" | "settings";
export type SettingsPanelSection = "general" | "conns" | "ai" | "persona" | "signature" | "references" | "groups" | "crm2layout" | "protection";
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

export type EmailIngestionTone = "red" | "orange" | "green";

export interface EmailIngestionStatus {
    identity: string;
    tone: EmailIngestionTone;
    detail: string;
    progress: number;
    isRunning: boolean;
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

interface ConnectivitySnapshot {
    connectionStatus: "none" | "success" | "error";
    granularStatus: { odoo: boolean | null; openai: boolean | null; gemini: boolean | null };
    granularStatusDetails: GranularStatusDetails;
    granularStatusString: string;
    checkedAt: number;
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
    settingsSection: SettingsPanelSection;
    setSettingsSection: (section: SettingsPanelSection) => void;
    openSettingsSection: (section: SettingsPanelSection) => void;
    emailIngestionStatus: EmailIngestionStatus;
}

// Export the context so it can be checked or used elsewhere if needed (rare)
// Using a global singleton pattern to prevent duplication in Outlook/Vite HMR
type CockpitContextSingletonHost = typeof globalThis & {
    __ICCC_COCKPIT_CONTEXT_v1__?: React.Context<CockpitContextType | undefined>;
};

const G = globalThis as CockpitContextSingletonHost;
const GK = "__ICCC_COCKPIT_CONTEXT_v1__";
const ACTIVE_TAB_STORAGE_KEY = "iccc_active_tab_v1";
const ACTIVE_SETTINGS_SECTION_STORAGE_KEY = "iccc_settings_section_v1";
const CONNECTIVITY_CACHE_STORAGE_KEY = "iccc_connectivity_status_v1";
const AI_CACHE_STORAGE_KEY = "iccc_ai_cache_v1";
const AI_CACHE_LEGACY_STORAGE_KEY = "icc_ai_cache_v1";
const WARM_BOOT_STORAGE_KEY = "iccc_warm_boot_v1";
const WARM_BOOT_MAX_AGE_MS = 10 * 60 * 1000;
const LINKS_CACHE_PREFIX = "iccc_links_cache_v1:";
const LINKS_CACHE_MESSAGE_PREFIX = "iccc_links_cache_msg_v1:";
const LINKS_CACHE_ITEM_PREFIX = "iccc_links_cache_item_v1:";
const AI_CACHE_ENTRY_TTL_MS = 5 * 24 * 60 * 60 * 1000;
const AI_CACHE_MAX_CONVERSATIONS = 6;
const AI_CACHE_MAX_PROMPT_CHARS = 2000;
const AI_CACHE_MAX_OUTPUT_CHARS = 6000;
const AI_CACHE_MAX_HISTORY_ITEMS = 12;
const AI_CACHE_MAX_HISTORY_CHARS = 1000;
const AI_CACHE_MAX_SMART_REPLIES = 6;
const AI_CACHE_MAX_SMART_REPLY_CHARS = 300;
const AI_CACHE_MAX_RECIPIENTS = 8;
const AI_CACHE_MAX_SUBJECT_CHARS = 500;
const AI_CACHE_MAX_TOTAL_CHARS = 120_000;
type PersistedAiCacheEntry = {
    state: AiState;
    updatedAtMs: number;
};
type PersistedAiCache = Record<string, PersistedAiCacheEntry>;
let aiCacheQuotaWarningLogged = false;
const INITIAL_EMAIL_INGESTION_STATUS: EmailIngestionStatus = {
    identity: "",
    tone: "red",
    detail: "Nenhum email persistido ainda.",
    progress: 100,
    isRunning: false,
};
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

function createEmptyAiState(): AiState {
    return {
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
    };
}

function trimPersistedText(value: string | undefined, maxChars: number): string {
    const raw = String(value || "").trim();
    if (!raw) return "";
    if (raw.length <= maxChars) return raw;
    return raw.slice(0, maxChars).trimEnd();
}

function trimPersistedStringArray(values: string[] | undefined, maxItems: number, maxChars: number): string[] {
    return Array.isArray(values)
        ? values
            .map((value) => trimPersistedText(value, maxChars))
            .filter(Boolean)
            .slice(-maxItems)
        : [];
}

function normalizePersistedAiState(state: Partial<AiState> | null | undefined, compact = false): AiState {
    const historyItemLimit = compact ? Math.min(6, AI_CACHE_MAX_HISTORY_ITEMS) : AI_CACHE_MAX_HISTORY_ITEMS;
    const historyCharLimit = compact ? Math.min(400, AI_CACHE_MAX_HISTORY_CHARS) : AI_CACHE_MAX_HISTORY_CHARS;
    const smartReplyLimit = compact ? Math.min(3, AI_CACHE_MAX_SMART_REPLIES) : AI_CACHE_MAX_SMART_REPLIES;
    const smartReplyCharLimit = compact ? Math.min(180, AI_CACHE_MAX_SMART_REPLY_CHARS) : AI_CACHE_MAX_SMART_REPLY_CHARS;
    const promptLimit = compact ? Math.min(1000, AI_CACHE_MAX_PROMPT_CHARS) : AI_CACHE_MAX_PROMPT_CHARS;
    const outputLimit = compact ? Math.min(2500, AI_CACHE_MAX_OUTPUT_CHARS) : AI_CACHE_MAX_OUTPUT_CHARS;

    return {
        prompt: trimPersistedText(state?.prompt, promptLimit),
        output: trimPersistedText(state?.output, outputLimit),
        tone: state?.tone || "neutro",
        locale: state?.locale || "auto",
        history: Array.isArray(state?.history)
            ? state.history
                .filter((entry) => entry && (entry.role === "user" || entry.role === "assistant"))
                .slice(-historyItemLimit)
                .map((entry) => ({
                    role: entry.role,
                    content: trimPersistedText(entry.content, historyCharLimit),
                }))
                .filter((entry) => entry.content)
            : [],
        smartReplies: trimPersistedStringArray(state?.smartReplies, smartReplyLimit, smartReplyCharLimit),
        action: trimPersistedText(state?.action, 80) || "reply",
        suggestedTo: trimPersistedStringArray(state?.suggestedTo, AI_CACHE_MAX_RECIPIENTS, 160),
        suggestedCc: trimPersistedStringArray(state?.suggestedCc, AI_CACHE_MAX_RECIPIENTS, 160),
        suggestedSubject: trimPersistedText(state?.suggestedSubject, AI_CACHE_MAX_SUBJECT_CHARS),
    };
}

function hasUsefulAiState(state: AiState): boolean {
    return Boolean(
        state.prompt ||
        state.output ||
        state.history.length ||
        state.smartReplies.length ||
        state.suggestedTo?.length ||
        state.suggestedCc?.length ||
        state.suggestedSubject
    );
}

function coerceAiCacheTimestamp(value: unknown): number | null {
    if (typeof value === "number" && Number.isFinite(value) && value > 0) return value;
    if (typeof value === "string") {
        const numeric = Number(value);
        if (Number.isFinite(numeric) && numeric > 0) return numeric;
        const parsed = Date.parse(value);
        if (Number.isFinite(parsed) && parsed > 0) return parsed;
    }
    return null;
}

function normalizeAiCacheForPersistence(
    cache: Record<string, AiState | PersistedAiCacheEntry> | null | undefined,
    compact = false
): PersistedAiCache {
    const now = Date.now();
    const entries = Object.entries(cache || {})
        .map(([conversationId, rawEntry]) => {
            const normalizedConversationId = String(conversationId || "").trim();
            if (!normalizedConversationId || !rawEntry || typeof rawEntry !== "object") return null;
            const rawObject = rawEntry as Partial<PersistedAiCacheEntry> & Partial<AiState> & {
                cachedAtMs?: unknown;
                createdAtMs?: unknown;
                updatedAt?: unknown;
                createdAt?: unknown;
            };
            const hasWrappedState = rawObject.state && typeof rawObject.state === "object";
            const rawState = hasWrappedState ? rawObject.state : rawObject;
            const normalizedState = normalizePersistedAiState(rawState, compact);
            if (!hasUsefulAiState(normalizedState)) return null;
            const timestamp = coerceAiCacheTimestamp(rawObject.updatedAtMs)
                ?? coerceAiCacheTimestamp(rawObject.cachedAtMs)
                ?? coerceAiCacheTimestamp(rawObject.createdAtMs)
                ?? coerceAiCacheTimestamp(rawObject.updatedAt)
                ?? coerceAiCacheTimestamp(rawObject.createdAt)
                ?? now;
            if (now - timestamp > AI_CACHE_ENTRY_TTL_MS) return null;
            return [normalizedConversationId, { state: normalizedState, updatedAtMs: timestamp }] as const;
        })
        .filter((entry): entry is readonly [string, PersistedAiCacheEntry] => Boolean(entry))
        .sort((a, b) => a[1].updatedAtMs - b[1].updatedAtMs)
        .slice(-AI_CACHE_MAX_CONVERSATIONS);
    return Object.fromEntries(entries);
}

function serializeAiCacheForPersistence(cache: PersistedAiCache): string {
    const fullSnapshot = normalizeAiCacheForPersistence(cache, false);
    const fullJson = JSON.stringify(fullSnapshot);
    if (fullJson.length <= AI_CACHE_MAX_TOTAL_CHARS) return fullJson;

    const compactSnapshot = normalizeAiCacheForPersistence(cache, true);
    const compactEntries = Object.entries(compactSnapshot);
    while (compactEntries.length > 1) {
        const nextJson = JSON.stringify(Object.fromEntries(compactEntries));
        if (nextJson.length <= AI_CACHE_MAX_TOTAL_CHARS) return nextJson;
        compactEntries.shift();
    }

    return JSON.stringify(Object.fromEntries(compactEntries));
}

function readPersistedAiCache(): PersistedAiCache {
    try {
        const saved = localStorage.getItem(AI_CACHE_STORAGE_KEY)
            || localStorage.getItem(AI_CACHE_LEGACY_STORAGE_KEY);
        if (!saved) return {};
        const parsed = JSON.parse(saved);
        if (!parsed || typeof parsed !== "object") return {};
        return normalizeAiCacheForPersistence(parsed as Record<string, AiState>);
    } catch {
        return {};
    }
}

function isQuotaExceededStorageError(error: unknown): boolean {
    if (!error || typeof error !== "object") return false;
    const candidate = error as { name?: string; code?: number; message?: string };
    return candidate.name === "QuotaExceededError"
        || candidate.code === 22
        || /quota/i.test(String(candidate.message || ""));
}

function persistAiCacheSnapshot(cache: PersistedAiCache): PersistedAiCache {
    const normalizedCache = normalizeAiCacheForPersistence(cache);
    if (!Object.keys(normalizedCache).length) {
        localStorage.removeItem(AI_CACHE_STORAGE_KEY);
        localStorage.removeItem(AI_CACHE_LEGACY_STORAGE_KEY);
        return {};
    }

    try {
        localStorage.setItem(AI_CACHE_STORAGE_KEY, serializeAiCacheForPersistence(normalizedCache));
        localStorage.removeItem(AI_CACHE_LEGACY_STORAGE_KEY);
        return normalizedCache;
    } catch (error) {
        if (!isQuotaExceededStorageError(error)) throw error;
    }

    const compactCache = normalizeAiCacheForPersistence(cache, true);
    if (!Object.keys(compactCache).length) {
        localStorage.removeItem(AI_CACHE_STORAGE_KEY);
        localStorage.removeItem(AI_CACHE_LEGACY_STORAGE_KEY);
        return {};
    }

    try {
        localStorage.setItem(AI_CACHE_STORAGE_KEY, serializeAiCacheForPersistence(compactCache));
        localStorage.removeItem(AI_CACHE_LEGACY_STORAGE_KEY);
        return compactCache;
    } catch (error) {
        if (!isQuotaExceededStorageError(error)) throw error;
        localStorage.removeItem(AI_CACHE_STORAGE_KEY);
        localStorage.removeItem(AI_CACHE_LEGACY_STORAGE_KEY);
        if (!aiCacheQuotaWarningLogged) {
            aiCacheQuotaWarningLogged = true;
            clientLog("warn", "[Cockpit] AI cache exceeded localStorage quota and was cleared", error);
        }
        return {};
    }
}

function upsertAiCacheEntry(cache: PersistedAiCache, conversationId: string, nextState: AiState): PersistedAiCache {
    const normalizedConversationId = String(conversationId || "").trim();
    if (!normalizedConversationId) return cache;
    const normalizedState = normalizePersistedAiState(nextState);
    const nextCache = { ...(cache || {}) };
    if (!hasUsefulAiState(normalizedState)) {
        delete nextCache[normalizedConversationId];
        return normalizeAiCacheForPersistence(nextCache);
    }
    nextCache[normalizedConversationId] = { state: normalizedState, updatedAtMs: Date.now() };
    return normalizeAiCacheForPersistence(nextCache);
}

function isSettingsPanelSection(value: string | null): value is SettingsPanelSection {
    return value === "general" ||
        value === "conns" ||
        value === "ai" ||
        value === "persona" ||
        value === "signature" ||
        value === "references" ||
        value === "crm2layout" ||
        value === "groups" ||
        value === "protection";
}

function buildContextEmailKey(ctx: OutlookMessageContext): string {
    return [
        String(ctx.itemId || "").trim(),
        String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, ""),
        String(ctx.conversationId || "").trim(),
    ].join("|");
}

function buildOutlookCategorySyncIdentity(ctx: OutlookMessageContext): string {
    const internetMessageId = String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
    const itemId = String(ctx.itemId || "").trim();
    if (internetMessageId) return `imid:${internetMessageId}`;
    if (itemId) return `item:${itemId}`;
    return "";
}

function doesRelatedEmailMatchContext(email: {
    itemId?: string;
    internetMessageId?: string;
} | null | undefined, ctx: OutlookMessageContext): boolean {
    const currentItemId = String(ctx.itemId || "").trim();
    const currentInternetMessageId = String(ctx.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
    const emailItemId = String(email?.itemId || "").trim();
    const emailInternetMessageId = String(email?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
    if (currentItemId) return Boolean(emailItemId) && emailItemId === currentItemId;
    if (currentInternetMessageId) return Boolean(emailInternetMessageId) && emailInternetMessageId === currentInternetMessageId;
    return false;
}

function readPersistedConnectivitySnapshot(): ConnectivitySnapshot | null {
    try {
        const raw = localStorage.getItem(CONNECTIVITY_CACHE_STORAGE_KEY);
        if (!raw) return null;
        const parsed = JSON.parse(raw);
        if (!parsed || typeof parsed !== "object") return null;
        const status = parsed.connectionStatus;
        if (status !== "none" && status !== "success" && status !== "error") return null;
        return {
            connectionStatus: status,
            granularStatus: {
                odoo: typeof parsed.granularStatus?.odoo === "boolean" ? parsed.granularStatus.odoo : null,
                openai: typeof parsed.granularStatus?.openai === "boolean" ? parsed.granularStatus.openai : null,
                gemini: typeof parsed.granularStatus?.gemini === "boolean" ? parsed.granularStatus.gemini : null,
            },
            granularStatusDetails: {
                openai: typeof parsed.granularStatusDetails?.openai === "string" ? parsed.granularStatusDetails.openai : null,
                gemini: typeof parsed.granularStatusDetails?.gemini === "string" ? parsed.granularStatusDetails.gemini : null,
                geminiDetails: parsed.granularStatusDetails?.geminiDetails || null,
            },
            granularStatusString: typeof parsed.granularStatusString === "string"
                ? parsed.granularStatusString
                : "Odoo: -- | OpenAI: -- | Gemini: --",
            checkedAt: Number(parsed.checkedAt || 0),
        };
    } catch {
        return null;
    }
}

function persistConnectivitySnapshot(snapshot: ConnectivitySnapshot) {
    try {
        localStorage.setItem(CONNECTIVITY_CACHE_STORAGE_KEY, JSON.stringify(snapshot));
    } catch {
        // ignore persistence failures
    }
}

if (!G[GK]) {
    G[GK] = createContext<CockpitContextType | undefined>(undefined);
}
export const CockpitContext = G[GK] as React.Context<CockpitContextType | undefined>;

export const CockpitProvider: React.FC<{ children: React.ReactNode }> = ({ children }) => {
    const initialConnectivitySnapshot = readPersistedConnectivitySnapshot();

    function readPersistedTab(): CockpitTab {
        try {
            const raw = sessionStorage.getItem(ACTIVE_TAB_STORAGE_KEY);
            return raw === "ai" || raw === "crm" || raw === "crm2" || raw === "related" || raw === "groups" || raw === "files"
                ? raw
                : "ai";
        } catch {
            return "ai";
        }
    }

    function readPersistedSettingsSection(): SettingsPanelSection {
        try {
            const raw = sessionStorage.getItem(ACTIVE_SETTINGS_SECTION_STORAGE_KEY);
            return isSettingsPanelSection(raw) ? raw : "general";
        } catch {
            return "general";
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
    const currentViewRef = useRef<string>((new URLSearchParams(window.location.search).get("view") || "taskpane").toLowerCase());
    const liveItemTrackingEnabled = currentViewRef.current !== "group-classification-studio";
    const [tab, setTab] = useState<CockpitTab>(() => readPersistedTab());
    const [settingsSection, setSettingsSectionState] = useState<SettingsPanelSection>(() => readPersistedSettingsSection());
    const [ctx, setCtx] = useState<OutlookMessageContext>({});
    const [bodyText, setBodyText] = useState<string>("");
    const [bodyHtml, setBodyHtml] = useState<string>("");
    const [meta, setMeta] = useState<OdooMeta | null>(null);
    const [links, setLinks] = useState<LinkEntry[]>([]);
    const [attachments, setAttachments] = useState<OutlookAttachment[]>([]);
    const [msg, setMsg] = useState<string | null>(null);
    const [isLoading, setIsLoading] = useState<boolean>(() => !warmStartupRef.current);
    const [isAuthenticated, setIsAuthenticated] = useState<boolean>(() => Boolean(getCachedSettingsSnapshot()?.odooSessionToken));
    const [connectionStatus, setConnectionStatus] = useState<"none" | "success" | "error">(
        () => initialConnectivitySnapshot?.connectionStatus || "none"
    );
    const [settings, setSettings] = useState<CockpitSettingsV1 | null>(() => getCachedSettingsSnapshot());
    const [activeGroupSelection, setActiveGroupSelection] = useState<{ emailKey: string; groupId: string | null }>({
        emailKey: "",
        groupId: null,
    });
    const [currentOutlookCategorySource, setCurrentOutlookCategorySource] = useState<OutlookCategorySource>({
        principalGroupNames: [],
        referenceGroupNames: [],
        ticketCodes: [],
        labelNames: [],
        managedLabelNames: [],
        groupStatuses: [],
        ticketStatuses: [],
        labelStatuses: [],
        specialCategories: [],
        managedSpecialCategories: [],
    });
    const [startupChecks, setStartupChecks] = useState<StartupCheck[]>(() => createStartupChecks());
    const [startupNoticeState, setStartupNoticeState] = useState<StartupNotice | null>(null);
    const [startupNoticeDismissed, setStartupNoticeDismissed] = useState(false);
    const [outlookCategoryRefreshTick, setOutlookCategoryRefreshTick] = useState(0);
    const [currentOutlookCategorySourceStatus, setCurrentOutlookCategorySourceStatus] = useState<{
        identity: string;
        ready: boolean;
    }>({
        identity: "",
        ready: false,
    });
    const [emailIngestionStatus, setEmailIngestionStatus] = useState<EmailIngestionStatus>(INITIAL_EMAIL_INGESTION_STATUS);
    const emailIngestionStatusRef = useRef<EmailIngestionStatus>(INITIAL_EMAIL_INGESTION_STATUS);
    const lastOutlookCategorySyncRef = useRef<{ identity: string; signature: string } | null>(null);
    const processedOutlookCategorySyncRequestRef = useRef("");
    const [outlookCategorySyncRequestTick, setOutlookCategorySyncRequestTick] = useState(0);

    function commitEmailIngestionStatus(next: EmailIngestionStatus) {
        emailIngestionStatusRef.current = next;
        setEmailIngestionStatus(next);
    }

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

    const setActiveGroupForCurrentEmail = useCallback((groupId: string | null) => {
        const emailKey = buildContextEmailKey(ctx);
        const nextGroupId = groupId ? String(groupId).trim() : null;
        setActiveGroupSelection((current) => {
            if (current.emailKey === emailKey && current.groupId === nextGroupId) {
                return current;
            }
            return { emailKey, groupId: nextGroupId };
        });
    }, [ctx]);

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
        const handleOutlookCategoryContextInvalidated = () => {
            setOutlookCategoryRefreshTick((current) => current + 1);
        };
        window.addEventListener(OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT, handleOutlookCategoryContextInvalidated as EventListener);
        return () => window.removeEventListener(OUTLOOK_CATEGORY_CONTEXT_INVALIDATED_EVENT, handleOutlookCategoryContextInvalidated as EventListener);
    }, []);

    useEffect(() => {
        const notify = () => setOutlookCategorySyncRequestTick((current) => current + 1);
        const handleStorage = (event: StorageEvent) => {
            if (event.key && event.key !== OUTLOOK_CATEGORY_SYNC_REQUEST_STORAGE_KEY) return;
            notify();
        };
        window.addEventListener("storage", handleStorage as EventListener);
        window.addEventListener(OUTLOOK_CATEGORY_SYNC_REQUEST_EVENT, notify as EventListener);
        return () => {
            window.removeEventListener("storage", handleStorage as EventListener);
            window.removeEventListener(OUTLOOK_CATEGORY_SYNC_REQUEST_EVENT, notify as EventListener);
        };
    }, []);

    useEffect(() => {
        let cancelled = false;
        void (async () => {
            const pendingRequest = readPendingOutlookCategorySyncRequest();
            if (!pendingRequest?.requestId) return;
            if (processedOutlookCategorySyncRequestRef.current === pendingRequest.requestId) return;
            if (pendingRequest.target && !doesRelatedEmailMatchContext(pendingRequest.target, ctx)) return;

            const expectedItemToken = await getCurrentItemToken().catch(() => "");
            if (cancelled) return;

            try {
                const shouldConsumeRequestResult = (result: { result: string }) =>
                    result.result !== "item-mismatch" && result.result !== "timeout";

                if (pendingRequest.mode === "source" && pendingRequest.source) {
                    const result = await executeOutlookCategorySourceSync(pendingRequest.source, {
                        expectedItemToken,
                        requestId: pendingRequest.requestId,
                        operationId: pendingRequest.operationId,
                        requestedAtIso: pendingRequest.createdAtIso,
                        reason: pendingRequest.reason || "taskpane-fallback-source",
                        target: pendingRequest.target,
                    }).catch(() => null);
                    if (cancelled || !result) return;
                    if (shouldConsumeRequestResult(result)) {
                        processedOutlookCategorySyncRequestRef.current = pendingRequest.requestId;
                    }
                    if (result.result === "success" || result.result === "duplicate") {
                        lastOutlookCategorySyncRef.current = null;
                        setOutlookCategoryRefreshTick((current) => current + 1);
                    }
                    return;
                }

                const result = await executeCurrentItemOutlookCategorySync({
                    expectedItemToken,
                    requestId: pendingRequest.requestId,
                    operationId: pendingRequest.operationId,
                    requestedAtIso: pendingRequest.createdAtIso,
                    reason: pendingRequest.reason || "taskpane-fallback-current-item-context",
                    target: pendingRequest.target,
                }).catch(() => null);
                if (cancelled || !result) return;
                if (shouldConsumeRequestResult(result)) {
                    processedOutlookCategorySyncRequestRef.current = pendingRequest.requestId;
                }
                if (result.result === "success" || result.result === "duplicate") {
                    lastOutlookCategorySyncRef.current = null;
                    setOutlookCategoryRefreshTick((current) => current + 1);
                }
            } catch (error) {
                clientLog.warn("[Cockpit] pending outlook category sync request failed", error);
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [ctx.internetMessageId, ctx.itemId, outlookCategorySyncRequestTick]);

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

    useEffect(() => {
        try {
            sessionStorage.setItem(ACTIVE_SETTINGS_SECTION_STORAGE_KEY, settingsSection);
        } catch {
            // ignore storage failures
        }
    }, [settingsSection]);

    const setSettingsSection = useCallback((section: SettingsPanelSection) => {
        setSettingsSectionState(section);
    }, []);

    const openSettingsSection = useCallback((section: SettingsPanelSection) => {
        setSettingsSectionState(section);
        if (currentViewRef.current === "app-settings") return;
        void openAppSettings({ section });
    }, []);

    // AI History Persistence
    const aiCacheRef = useRef<PersistedAiCache>({});
    const [aiCache, setAiCache] = useState<PersistedAiCache>(() => {
        const persisted = readPersistedAiCache();
        aiCacheRef.current = persisted;
        return persisted;
    });

    // Save AI cache whenever it changes
    useEffect(() => {
        try {
            const persisted = persistAiCacheSnapshot(aiCache);
            aiCacheRef.current = persisted;
            if (Object.keys(aiCache).length && !Object.keys(persisted).length) {
                setAiCache({});
            }
        } catch (e) {
            clientLog("error", "[Cockpit] Failed to save AI cache to localStorage", e);
        }
    }, [aiCache]);

    const [currentAiState, setCurrentAiState] = useState<AiState>(createEmptyAiState);

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
        let retryForIdentityStabilization = false;
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
                let authCheck: AuthCheckResponse = { ok: true, authenticated: false, reason: "startup_not_checked" };
                try {
                    authCheck = await apiCheckAuth();
                } catch (error) {
                    clientLog("warn", "[Cockpit] apiCheckAuth failed during startup", error);
                }
                if (authCheck.ok && authCheck.authenticated) {
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
            const stableSelection = await waitForStableSelectedMessageContext({
                maxAttempts: 4,
                delayMs: 120,
                requirePreciseIdentity: true,
            }).catch(() => ({
                context: {} as OutlookMessageContext,
                itemToken: "",
            }));
            const c: OutlookMessageContext = stableSelection.context || {};
            const ingestionIdentity = buildContextEmailKey(c);
            const currentIngestion = emailIngestionStatusRef.current;
            const isNewIngestion = ingestionIdentity !== currentIngestion.identity;
            if (ingestionIdentity) {
                if (isNewIngestion) {
                    commitEmailIngestionStatus({
                        identity: ingestionIdentity,
                        tone: "green",
                        detail: "A preparar a leitura do email atual.",
                        progress: 8,
                        isRunning: true,
                    });
                } else if (currentIngestion.isRunning) {
                    commitEmailIngestionStatus({
                        ...currentIngestion,
                        tone: "green",
                        detail: "A preparar a leitura do email atual.",
                        progress: Math.max(currentIngestion.progress, 8),
                        isRunning: true,
                    });
                } else if (currentIngestion.tone !== "green") {
                    commitEmailIngestionStatus({
                        ...currentIngestion,
                        tone: "green",
                        detail: "A retomar a leitura do email atual.",
                        progress: Math.max(currentIngestion.progress, 8),
                        isRunning: true,
                    });
                }
            }
            const hasPreciseContextIdentity = Boolean(String(c.itemId || "").trim() || String(c.internetMessageId || "").trim());
            if (!hasPreciseContextIdentity) {
                retryForIdentityStabilization = true;
            }
            if (ingestionIdentity) {
                const nextIngestion = emailIngestionStatusRef.current;
                if (nextIngestion.identity === ingestionIdentity && nextIngestion.isRunning) {
                    commitEmailIngestionStatus({
                        ...nextIngestion,
                        detail: hasPreciseContextIdentity
                            ? "Identidade confirmada. A ler corpo e anexos."
                            : "A identidade do email ainda nao esta estavel.",
                        progress: Math.max(nextIngestion.progress, hasPreciseContextIdentity ? 22 : 16),
                        isRunning: true,
                    });
                }
            }
            const emailLoadLabels = ["body-text", "body-html", "attachments"] as const;
            const emailLoadResults = hasPreciseContextIdentity
                ? await Promise.allSettled([
                    getEmailBodyText(),
                    getEmailBodyHtml(),
                    import("@/office").then((m) => m.getAttachments()),
                ])
                : [];
            const b = emailLoadResults[0]?.status === "fulfilled" ? emailLoadResults[0].value : "";
            const bh = emailLoadResults[1]?.status === "fulfilled" ? emailLoadResults[1].value : "";
            const atts: OutlookAttachment[] = emailLoadResults[2]?.status === "fulfilled" ? emailLoadResults[2].value : [];
            const emailLoadFailures = emailLoadResults.flatMap((result, index) => (
                result.status === "rejected" ? [emailLoadLabels[index]] : []
            ));

            if (reqId !== ctxLoadSeqRef.current) return;

            // Update token tracker to avoid redundant polls
            lastItemTokenRef.current = stableSelection.itemToken || await getCurrentItemToken().catch(() => "");
            setCtx(c);
            setBodyText(b);
            setBodyHtml(bh);
            setAttachments(atts || []);
            if (!hasPreciseContextIdentity) {
                commitEmailIngestionStatus({
                    identity: ingestionIdentity,
                    tone: "red",
                    detail: "A identidade do email ainda nao esta estavel. Nada foi enviado para a base de dados.",
                    progress: 100,
                    isRunning: false,
                });
            }
            if (emailLoadFailures.length) {
                clientLog("warn", "[Cockpit] partial email bootstrap load", { failures: emailLoadFailures });
            }
            if (hasPreciseContextIdentity && ingestionIdentity) {
                const nextIngestion = emailIngestionStatusRef.current;
                if (nextIngestion.identity === ingestionIdentity && nextIngestion.isRunning) {
                    commitEmailIngestionStatus({
                        ...nextIngestion,
                        tone: emailLoadFailures.length ? "orange" : "green",
                        detail: emailLoadFailures.length
                            ? "Leitura do email com falhas parciais. A continuar a gravacao."
                            : "Corpo e anexos lidos. A gravar o email na base de dados.",
                        progress: Math.max(nextIngestion.progress, 62),
                        isRunning: true,
                    });
                }
            }
            if (reason === "init") {
                const hasEmailContext = Boolean(c.itemId || c.internetMessageId || c.conversationId || c.subject);
                if (hasEmailContext && hasPreciseContextIdentity && !emailLoadFailures.length) {
                    updateStartupCheck("email", {
                        status: "success",
                        detail: c.subject
                            ? `Email pronto: ${c.subject}`
                            : "Contexto do email atual carregado.",
                    });
                } else if (!hasPreciseContextIdentity) {
                    updateStartupCheck("email", {
                        status: "warning",
                        detail: "O Outlook ainda esta a estabilizar a identidade do email aberto. Vamos revalidar antes de continuar.",
                    });
                } else if (hasEmailContext) {
                    updateStartupCheck("email", {
                        status: "warning",
                        detail: c.subject
                            ? `Email carregado com dados parciais: ${c.subject}`
                            : "O email abriu com alguns dados ainda por sincronizar.",
                    });
                } else if (emailLoadFailures.length) {
                    updateStartupCheck("email", {
                        status: "warning",
                        detail: "O Outlook ainda não disponibilizou o email completo. Vamos continuar e tentar novamente em segundo plano.",
                    });
                } else {
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
            if (!hasPreciseContextIdentity) {
                setLinks([]);
                return;
            }
            if (ingestionIdentity) {
                const nextIngestion = emailIngestionStatusRef.current;
                if (nextIngestion.identity === ingestionIdentity) {
                    commitEmailIngestionStatus({
                        ...nextIngestion,
                        tone: nextIngestion.tone === "red" ? "orange" : nextIngestion.tone,
                        detail: "A gravar o email atual e os anexos na base de dados.",
                        progress: Math.max(nextIngestion.progress, 84),
                        isRunning: true,
                    });
                }
            }
            const shouldRegisterEmailRemotely = tab !== "groups";
            let emailRegistryOk = !shouldRegisterEmailRemotely;
            if (shouldRegisterEmailRemotely) {
                try {
                    const attachmentStorage = getGroupAttachmentStorageOptions(s);
                    await registerRelevantEmail({
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
                        ...attachmentStorage,
                        attachments: (atts || []).map((attachment) => ({
                            key: String((attachment as any)?.key || "").trim() || undefined,
                            name: attachment.name,
                            contentType: attachment.contentType,
                            size: attachment.size,
                            id: attachment.id,
                            isInline: attachment.isInline,
                            contentId: attachment.contentId,
                            content: String(attachment.content || "").trim(),
                        })),
                    });
                    emailRegistryOk = true;
                } catch (error) {
                    clientLog("warn", "[Cockpit] email registry failed", error);
                }
            }
            if (reqId !== ctxLoadSeqRef.current) return;

            const attachmentCount = Array.isArray(atts) ? atts.length : 0;
            const attachmentLoadFailed = emailLoadFailures.includes("attachments");
            const bodyLoadFailed = emailLoadFailures.includes("body-text") || emailLoadFailures.includes("body-html");
            if (!shouldRegisterEmailRemotely) {
                commitEmailIngestionStatus({
                    identity: ingestionIdentity,
                    tone: emailLoadFailures.length ? "orange" : "green",
                    detail: emailLoadFailures.length
                        ? "Preparar manteve o email em sessao/local com dados parciais."
                        : "Preparar manteve o email em sessao/local, sem envio automatico ao servidor.",
                    progress: 100,
                    isRunning: false,
                });
            } else if (!emailRegistryOk) {
                commitEmailIngestionStatus({
                    identity: ingestionIdentity,
                    tone: "red",
                    detail: "Falhou o envio do email atual para a base de dados.",
                    progress: 100,
                    isRunning: false,
                });
            } else if (emailLoadFailures.length) {
                const partialBits = [
                    bodyLoadFailed ? "corpo parcial" : "",
                    attachmentLoadFailed ? `anexos parciais (${attachmentCount})` : "",
                ].filter(Boolean).join(" · ");
                commitEmailIngestionStatus({
                    identity: ingestionIdentity,
                    tone: "orange",
                    detail: partialBits
                        ? `Email enviado com dados parciais: ${partialBits}.`
                        : "Email enviado com dados parciais.",
                    progress: 100,
                    isRunning: false,
                });
            } else {
                commitEmailIngestionStatus({
                    identity: ingestionIdentity,
                    tone: "green",
                    detail: attachmentCount > 0
                        ? `Email e ${attachmentCount} anexo(s) enviados para a base de dados.`
                        : "Email enviado para a base de dados.",
                    progress: 100,
                    isRunning: false,
                });
            }

            if (c.conversationId) {
                const cached = aiCacheRef.current[c.conversationId];
                setCurrentAiState(cached ? normalizePersistedAiState(cached.state) : createEmptyAiState());
            } else {
                setCurrentAiState(createEmptyAiState());
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
                    status: connectionStatus === "error" ? "warning" : "success",
                    detail: connectionStatus === "none"
                        ? (s.aiManualOnly === false
                            ? "Monitorização automática ativa para Odoo e IA."
                            : "Verificação manual disponível em Settings > Ligações.")
                        : `Último teste: ${granularStatusString}`,
                });
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
            if (retryForIdentityStabilization && reqId === ctxLoadSeqRef.current) {
                window.setTimeout(() => {
                    void loadContextAndLinks("identity-stabilize");
                }, 250);
            }
        }
    }

    const [granularStatus, setGranularStatus] = useState<{ odoo: boolean | null; openai: boolean | null; gemini: boolean | null }>(
        () => initialConnectivitySnapshot?.granularStatus || {
            odoo: null,
            openai: null,
            gemini: null
        }
    );

    const [granularStatusDetails, setGranularStatusDetails] = useState<GranularStatusDetails>(
        () => initialConnectivitySnapshot?.granularStatusDetails || {
            openai: null,
            gemini: null,
            geminiDetails: null
        }
    );

    const [granularStatusString, setGranularStatusString] = useState<string>(
        () => initialConnectivitySnapshot?.granularStatusString || "Odoo: -- | OpenAI: -- | Gemini: --"
    );

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
            persistConnectivitySnapshot({
                connectionStatus: finalOdooOk && a.ok ? "success" : "error",
                granularStatus: newStatus,
                granularStatusDetails: {
                    openai: a.openai?.error || (a.openai?.ok ? `Model: ${a.openai.modelUsed || 'default'}` : null),
                    gemini: a.gemini?.error || (a.gemini?.ok ? `Model: ${a.gemini.modelUsed || 'default'}` : null),
                    geminiDetails: a.gemini?.ok ? {
                        requested: a.gemini.requestedModel,
                        sanitized: a.gemini.sanitizedModel,
                        effective: a.gemini.modelUsed,
                        provider: a.gemini.providerUsed
                    } : null
                },
                granularStatusString: summary,
                checkedAt: Date.now(),
            });
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
            const failedStatus = {
                odoo: false,
                openai: false,
                gemini: false,
            };
            const failedDetails = {
                openai: "Falha no teste manual.",
                gemini: "Falha no teste manual.",
                geminiDetails: null,
            };
            setConnectionStatus("error");
            setGranularStatus(failedStatus);
            setGranularStatusDetails(failedDetails);
            setGranularStatusString("Odoo: Erro | OpenAI: Erro | Gemini: Erro");
            persistConnectivitySnapshot({
                connectionStatus: "error",
                granularStatus: failedStatus,
                granularStatusDetails: failedDetails,
                granularStatusString: "Odoo: Erro | OpenAI: Erro | Gemini: Erro",
                checkedAt: Date.now(),
            });
            return {
                odooOk: false,
                openaiOk: false,
                geminiOk: false,
                summary: "Odoo: Erro | OpenAI: Erro | Gemini: Erro",
                failures: ["Não foi possível concluir as verificações de conectividade."],
            };
        }
    }

    useEffect(() => {
        if (settings?.aiManualOnly !== false) return;

        void checkConnectivity();
        const intervalId = window.setInterval(() => {
            void checkConnectivity();
        }, 5 * 60 * 1000);

        return () => window.clearInterval(intervalId);
    }, [settings?.aiManualOnly]);

    useEffect(() => {
        loadContextAndLinks('init');

        if (!liveItemTrackingEnabled) {
            clientLog("info", `[Cockpit] live item tracking disabled for view ${currentViewRef.current}`);
            return;
        }

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
        const currentSourceIdentity = buildOutlookCategorySyncIdentity(ctx);
        const hasPreciseContextIdentity = Boolean(currentSourceIdentity);
        if (!hasPreciseContextIdentity) {
            const nextSource: OutlookCategorySource = {
                principalGroupNames: [],
                referenceGroupNames: [],
                ticketCodes: [],
                labelNames: [],
                managedLabelNames: [],
                groupStatuses: [],
                ticketStatuses: [],
                labelStatuses: [],
                specialCategories: links.length > 0 ? [ODOO_LINKED_CATEGORY] : [],
                managedSpecialCategories: [ODOO_LINKED_CATEGORY],
            };
            setCurrentOutlookCategorySource((current) => (
                areOutlookCategorySourcesEqual(current, nextSource) ? current : nextSource
            ));
            setCurrentOutlookCategorySourceStatus({
                identity: "",
                ready: false,
            });
            lastOutlookCategorySyncRef.current = null;
            return;
        }
        let cancelled = false;
        setCurrentOutlookCategorySourceStatus({
            identity: currentSourceIdentity,
            ready: false,
        });
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
                if (!doesRelatedEmailMatchContext(response?.email || null, ctx)) {
                    lastOutlookCategorySyncRef.current = null;
                    setCurrentOutlookCategorySourceStatus({
                        identity: currentSourceIdentity,
                        ready: false,
                    });
                    return;
                }
                const nextSource = buildOutlookCategorySourceFromRelatedContext({
                    email: response?.email || null,
                    groups: Array.isArray(response?.groups) ? response.groups : [],
                    tickets: Array.isArray(response?.tickets) ? response.tickets : [],
                    settings,
                    specialCategories: links.length > 0 ? [ODOO_LINKED_CATEGORY] : [],
                    managedSpecialCategories: [ODOO_LINKED_CATEGORY],
                });
                setCurrentOutlookCategorySource((current) => (
                    areOutlookCategorySourcesEqual(current, nextSource) ? current : nextSource
                ));
                setCurrentOutlookCategorySourceStatus({
                    identity: currentSourceIdentity,
                    ready: true,
                });
            })
            .catch(() => {
                if (!cancelled) {
                    const nextSource: OutlookCategorySource = {
                        principalGroupNames: [],
                        referenceGroupNames: [],
                        ticketCodes: [],
                        labelNames: [],
                        managedLabelNames: [],
                        groupStatuses: [],
                        ticketStatuses: [],
                        labelStatuses: [],
                        specialCategories: links.length > 0 ? [ODOO_LINKED_CATEGORY] : [],
                        managedSpecialCategories: [ODOO_LINKED_CATEGORY],
                    };
                    setCurrentOutlookCategorySource((current) => (
                        areOutlookCategorySourcesEqual(current, nextSource) ? current : nextSource
                    ));
                    lastOutlookCategorySyncRef.current = null;
                    setCurrentOutlookCategorySourceStatus({
                        identity: currentSourceIdentity,
                        ready: false,
                    });
                }
            });

        return () => {
            cancelled = true;
        };
    }, [ctx.conversationId, ctx.fromEmail, ctx.fromName, ctx.internetMessageId, ctx.itemId, ctx.receivedDateTimeIso, ctx.subject, links.length, outlookCategoryRefreshTick, settings?.groupOutlookCategories]);

    useEffect(() => {
        syncOdooLinkedNotification(links.length > 0, links.length).catch(() => {
            // best-effort host hint only
        });
        let cancelled = false;
        void (async () => {
            const syncIdentity = buildOutlookCategorySyncIdentity(ctx);
            if (!syncIdentity) return;
            if (
                currentOutlookCategorySourceStatus.identity !== syncIdentity
                || currentOutlookCategorySourceStatus.ready !== true
            ) {
                return;
            }
            const knownLabelNames = Array.from(new Set([
                ...currentOutlookCategorySource.managedLabelNames,
                ...(Array.isArray(settings?.groupLabelCatalog) ? settings.groupLabelCatalog.map((entry) => String(entry?.label || "").trim()).filter(Boolean) : []),
            ]));
            const snapshot = await getManagedOutlookCategorySnapshot(knownLabelNames).catch(() => null);
            if (cancelled) return;
            const expectedItemToken = await getCurrentItemToken().catch(() => "");
            if (cancelled) return;
            const syncSource = {
                ...currentOutlookCategorySource,
                managedLabelNames: [
                    ...currentOutlookCategorySource.managedLabelNames,
                    ...((snapshot?.labelNames || []).map((label) => String(label || "").trim()).filter(Boolean)),
                ],
            };
            const syncSignature = getOutlookCategoryPlanSignature(buildOutlookCategoryPlan(syncSource));
            const currentSnapshotSignature = snapshot
                ? getOutlookCategoryPlanSignature(buildOutlookCategoryPlan({
                    ...syncSource,
                    principalGroupNames: snapshot.principalGroupNames || [],
                    referenceGroupNames: snapshot.referenceGroupNames || [],
                    ticketCodes: snapshot.ticketCodes || [],
                    labelNames: snapshot.labelNames || [],
                    groupStatuses: snapshot.groupStatuses || [],
                    ticketStatuses: snapshot.ticketStatuses || [],
                    labelStatuses: snapshot.labelStatuses || [],
                }))
                : "";
            if (
                lastOutlookCategorySyncRef.current?.identity === syncIdentity
                && lastOutlookCategorySyncRef.current?.signature === syncSignature
                && currentSnapshotSignature === syncSignature
            ) {
                return;
            }
            const activeOperation = getActiveOutlookCategoryOperation({
                itemId: String(ctx.itemId || "").trim() || undefined,
                internetMessageId: String(ctx.internetMessageId || "").trim() || undefined,
                conversationId: String(ctx.conversationId || "").trim() || undefined,
            }, expectedItemToken);
            if (activeOperation) return;
            const result = await executeOutlookCategorySourceSync(syncSource, {
                expectedItemToken,
                reason: "provider-reconcile",
                target: {
                    itemId: String(ctx.itemId || "").trim() || undefined,
                    internetMessageId: String(ctx.internetMessageId || "").trim() || undefined,
                    conversationId: String(ctx.conversationId || "").trim() || undefined,
                },
            }).catch(() => null);
            if (!cancelled && (result?.result === "success" || result?.result === "duplicate")) {
                lastOutlookCategorySyncRef.current = {
                    identity: syncIdentity,
                    signature: syncSignature,
                };
            }
        })();
        return () => {
            cancelled = true;
        };
    }, [ctx.conversationId, ctx.internetMessageId, ctx.itemId, currentOutlookCategorySource, currentOutlookCategorySourceStatus, links.length, settings?.groupLabelCatalog]);

    const setAiState = (update: Partial<AiState>) => {
        if (!ctx.conversationId) return;
        setCurrentAiState(prev => {
            const newState = { ...prev, ...update };
            setAiCache((cache) => upsertAiCacheEntry(cache, ctx.conversationId!, newState));
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
            settingsSection,
            setSettingsSection,
            openSettingsSection,
            emailIngestionStatus,
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
