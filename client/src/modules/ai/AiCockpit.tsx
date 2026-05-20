import React, { useState, useEffect, useMemo, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { aiGenerate, type AiAction, type AiTone, type AiLocale, type AiReplyDirection, type AiSignaturePayload } from "@/ai/aiClient";
import { insertTextToBody, isComposeMode, displayReplyForm, displayForwardForm, displayNewMessageForm, displayNewMeetingForm, setRecipients, setSubjectInComposeDraft, openAiSettings, addBase64AttachmentToCompose, openAiReplyTargetPicker, syncLinkCategoriesToComposeDraft, type AiReplyTargetSelection } from "@/office";
import { getSettings, getSignatureImageDataUrl, type AppLocale, type CockpitSettingsV1 } from "@/settings";
import { getEmailAttachmentContentBase64, getRelatedEmailContext, logLearningInteraction, registerRelevantEmail, type RelevantEmailPayload, type RelatedEmailEntry } from "@/api";
import { getGroupAttachmentStorageOptions } from "@/modules/crm/groups-v1/storage/resolveStorageMode";
import { buildOutlookCategorySourceFromRelatedContext, mergeOutlookCategorySources, ODOO_LINKED_CATEGORY, type OutlookCategorySource } from "@/outlookCategories";
import { buildAiContextBundle, type AiContextBundle } from "./contextBundle";
import * as Icons from "@/ui/icons";

// --- PERSISTENCE HELPERS ---
const AI_HISTORY_KEY = "icc.ai_history.v1";
const HISTORY_KEEP_MS = 5 * 24 * 60 * 60 * 1000; // 5 days
const MAX_HISTORY_ENTRIES = 100;
const MAX_HISTORY_JSON_CHARS = 120_000;
const MAX_HISTORY_PROMPT_CHARS = 4_000;
const MAX_HISTORY_OUTPUT_CHARS = 12_000;

type HistoryEntry = {
    id: string;
    emailKey: string;
    ts: number;
    output: string;
    prompt: string;
    action: AiAction;
    tone: AiTone;
    locale: AiLocale;
    draftTo: string[];
    draftCc: string[];
    draftBcc?: string[];
    draftSubject: string;
    customToneId?: string;
    replyTarget?: AiReplyTargetSelection | null;
    replyDirection?: AiReplyDirection | null;
};

type QuickPanelId = "lang" | "mode" | "presets" | "intents" | "contacts" | "files" | null;

type FileUsageState = {
    analyze: boolean;
    forward: boolean;
};

type PersistedEmailAttachment = NonNullable<RelatedEmailEntry["attachments"]>[number];

function getEmailKey(ctx: any) {
    return ctx.conversationId || ctx.internetMessageId || "global";
}

function truncateHistoryText(value: unknown, maxChars: number): string {
    const raw = String(value || "").trim();
    return raw.length > maxChars ? raw.slice(0, maxChars).trimEnd() : raw;
}

function normalizeEmailListInput(value: unknown): string[] {
    const values = Array.isArray(value)
        ? value
        : String(value || "").split(/[;,]/);
    const seen = new Set<string>();
    const result: string[] = [];
    values.forEach((entry) => {
        const email = String(entry || "").trim();
        if (!email) return;
        const key = email.toLowerCase();
        if (seen.has(key)) return;
        seen.add(key);
        result.push(email);
    });
    return result;
}

function addUniqueEmail(values: string[], email: string): string[] {
    return normalizeEmailListInput([...values, email]);
}

function normalizeHistoryEntry(raw: any, now = Date.now()): HistoryEntry | null {
    if (!raw || typeof raw !== "object") return null;
    const ts = Number(raw.ts || 0);
    if (!Number.isFinite(ts) || ts <= 0 || now - ts >= HISTORY_KEEP_MS) return null;
    const output = truncateHistoryText(raw.output, MAX_HISTORY_OUTPUT_CHARS);
    const prompt = truncateHistoryText(raw.prompt, MAX_HISTORY_PROMPT_CHARS);
    if (!output && !prompt) return null;
    return {
        id: String(raw.id || `hist-${ts}-${Math.random().toString(36).slice(2, 8)}`),
        emailKey: String(raw.emailKey || "global"),
        ts,
        output,
        prompt,
        action: (raw.action || "reply") as AiAction,
        tone: (raw.tone || "neutro") as AiTone,
        locale: (raw.locale || "auto") as AiLocale,
        draftTo: normalizeEmailListInput(raw.draftTo),
        draftCc: normalizeEmailListInput(raw.draftCc),
        draftBcc: normalizeEmailListInput(raw.draftBcc),
        draftSubject: truncateHistoryText(raw.draftSubject, 300),
        customToneId: raw.customToneId ? String(raw.customToneId) : undefined,
        replyTarget: raw.replyTarget || null,
        replyDirection: raw.replyDirection || null,
    };
}

function writeHistory(entries: HistoryEntry[]): HistoryEntry[] {
    const normalized = entries
        .map((entry) => normalizeHistoryEntry(entry))
        .filter((entry): entry is HistoryEntry => Boolean(entry))
        .sort((a, b) => b.ts - a.ts)
        .slice(0, MAX_HISTORY_ENTRIES);
    while (normalized.length > 0 && JSON.stringify(normalized).length > MAX_HISTORY_JSON_CHARS) {
        normalized.pop();
    }
    try {
        if (!normalized.length) localStorage.removeItem(AI_HISTORY_KEY);
        else localStorage.setItem(AI_HISTORY_KEY, JSON.stringify(normalized));
    } catch {
        while (normalized.length > 1) {
            normalized.pop();
            try {
                localStorage.setItem(AI_HISTORY_KEY, JSON.stringify(normalized));
                break;
            } catch {
                // keep pruning oldest entries
            }
        }
    }
    return normalized;
}

function loadHistory(): HistoryEntry[] {
    try {
        const raw = localStorage.getItem(AI_HISTORY_KEY);
        if (!raw) return [];
        const arr = JSON.parse(raw);
        if (!Array.isArray(arr)) return [];
        const now = Date.now();
        const normalized = arr
            .map((entry: any) => normalizeHistoryEntry(entry, now))
            .filter((entry: HistoryEntry | null): entry is HistoryEntry => Boolean(entry))
            .sort((a: HistoryEntry, b: HistoryEntry) => b.ts - a.ts)
            .slice(0, MAX_HISTORY_ENTRIES);
        if (normalized.length !== arr.length || JSON.stringify(normalized) !== raw) {
            writeHistory(normalized);
        }
        return normalized;
    } catch { return []; }
}

function saveHistory(entries: HistoryEntry[]): HistoryEntry[] {
    return writeHistory(entries);
}

function htmlToPlainText(html: string): string {
    return String(html || "")
        .replace(/<style[\s\S]*?<\/style>/gi, " ")
        .replace(/<script[\s\S]*?<\/script>/gi, " ")
        .replace(/<br\s*\/?>/gi, "\n")
        .replace(/<\/p>/gi, "\n")
        .replace(/<[^>]+>/g, " ")
        .replace(/&nbsp;/gi, " ")
        .replace(/&amp;/gi, "&")
        .replace(/&lt;/gi, "<")
        .replace(/&gt;/gi, ">")
        .replace(/\r/g, "")
        .replace(/\n{3,}/g, "\n\n")
        .replace(/[ \t]{2,}/g, " ")
        .trim();
}

function escapeHtml(value: string): string {
    return String(value || "")
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/\"/g, "&quot;")
        .replace(/'/g, "&#39;");
}

function isRejectedAttachmentState(value: string | undefined): boolean {
    return String(value || "").trim().toLowerCase() === "rejected";
}

function looksLikeHtml(value: string): boolean {
    return /<\/?(p|br|ul|ol|li|strong|em|a|div|span|h[1-6]|table|tr|td|blockquote)\b/i.test(String(value || ""));
}

function htmlFragmentToPlainText(raw: string): string {
    const source = String(raw || "").trim();
    if (!source) return "";

    if (!looksLikeHtml(source)) {
        return source;
    }

    const marked = source
        .replace(/<\s*br\s*\/?\s*>/gi, "\n")
        .replace(/<\s*li[^>]*>/gi, "\n- ")
        .replace(/<\/\s*(p|div|li|h[1-6]|blockquote|tr|td)\s*>/gi, "\n\n");

    try {
        if (typeof document !== "undefined") {
            const div = document.createElement("div");
            div.innerHTML = marked;
            return div.textContent || "";
        }
    } catch {
        // fallback below
    }

    return marked.replace(/<[^>]+>/g, " ");
}

function normalizeGeneratedText(raw: string): string {
    return htmlFragmentToPlainText(raw)
        .replace(/\u00a0/g, " ")
        .replace(/\r\n/g, "\n")
        .replace(/\r/g, "\n")
        .replace(/\|\s*/g, "\n")
        .replace(/[ \t]+/g, " ")
        .replace(/[ \t]*\n[ \t]*/g, "\n")
        .replace(/\n{3,}/g, "\n\n")
        .trim();
}

function protectCommonAbbreviations(value: string): string {
    return String(value || "")
        .replace(/\b(Sr|Sra|Dr|Dra|Eng|Exmo|Exma|Mr|Mrs|Ms|St)\./gi, "$1§")
        .replace(/\be\.g\./gi, "e§g§")
        .replace(/\bi\.e\./gi, "i§e§");
}

function restoreCommonAbbreviations(value: string): string {
    return String(value || "")
        .replace(/§/g, ".");
}

function splitSentenceBlocks(line: string): string[] {
    const source = String(line || "").trim();
    if (!source) return [];

    const protectedText = protectCommonAbbreviations(source);
    const parts = protectedText.match(/[^.!?]+[.!?]+(?=\s|$)|[^.!?]+$/g) || [protectedText];

    return parts
        .map((part) => restoreCommonAbbreviations(part).trim())
        .filter(Boolean);
}

function forceEmailBoundaryBreaks(text: string): string {
    return String(text || "")
        .replace(/\s+(Com os melhores cumprimentos,|Melhores cumprimentos,|Cumprimentos,|Best regards,|Kind regards,|Regards,|Saludos,|Un saludo,|Cordialmente,|Cordiali saluti,|Mit freundlichen Grüßen,)/gi, "\n\n$1")
        .replace(/\s+(Muito obrigado(?: pela| pelo|,|\.|$)|Obrigado(?: pela| pelo|,|\.|$)|Thank you(?: for [^.!?]+[.!?]|\.|,|$)|Muchas gracias(?:[^.!?]*[.!?]|,|$)|Gracias(?:[^.!?]*[.!?]|,|$))/gi, "\n\n$1")
        .replace(/\s+(Pedro Lopes(?:\s+Backoffice Divitek)?)(?=\s|$)/gi, "\n\n$1");
}

function formatEmailHtml(raw: string): string {
    const source = normalizeGeneratedText(raw);
    if (!source) return "";

    const prepared = forceEmailBoundaryBreaks(source)
        .replace(/([A-Za-zÀ-ÿ]+,)\s+([A-ZÁÀÂÃÉÈÊÍÌÎÓÒÔÕÚÙÛÇ\[])/g, "$1\n\n$2")
        .replace(/\n{3,}/g, "\n\n")
        .trim();

    const blocks: string[] = [];
    let listItems: string[] = [];

    const flushList = () => {
        if (!listItems.length) return;
        blocks.push(
            `<ul style="margin:0 0 14px 18px;padding:0;">${
                listItems.map((item) => `<li style="margin:0 0 6px;">${escapeHtml(item)}</li>`).join("")
            }</ul>`
        );
        listItems = [];
    };

    const pushParagraph = (value: string) => {
        const clean = String(value || "").trim();
        if (!clean) return;
        flushList();
        blocks.push(`<p style="margin:0 0 14px;">${escapeHtml(clean)}</p>`);
    };

    const chunks = prepared.split(/\n{2,}/).map((chunk) => chunk.trim()).filter(Boolean);

    for (const chunk of chunks) {
        const lines = chunk.split("\n").map((line) => line.trim()).filter(Boolean);

        for (const rawLine of lines) {
            const line = String(rawLine || "").trim();
            if (!line) continue;

            if (/^[-*•]\s+/.test(line)) {
                listItems.push(line.replace(/^[-*•]\s+/, "").trim());
                continue;
            }

            flushList();

            const sentenceBlocks = splitSentenceBlocks(line);

            if (sentenceBlocks.length <= 1) {
                pushParagraph(line);
                continue;
            }

            for (const sentence of sentenceBlocks) {
                pushParagraph(sentence);
            }
        }
    }

    flushList();

    if (!blocks.length) {
        return "";
    }

    return `<div style="font-family:Aptos,Segoe UI,Arial,sans-serif;font-size:11pt;line-height:1.55;color:#1f2937;">${blocks.join("")}</div>`;
}

const AI_SIGNATURE_LOCALES: AppLocale[] = ["pt-PT", "es-ES", "en-GB", "it-IT", "de-DE"];

function resolveSignatureLocale(settings: CockpitSettingsV1, locale: AiLocale): AppLocale {
    if (locale !== "auto" && AI_SIGNATURE_LOCALES.includes(locale as AppLocale)) return locale as AppLocale;
    const replyLanguage = settings.replyLanguage;
    if (replyLanguage && replyLanguage !== "auto" && AI_SIGNATURE_LOCALES.includes(replyLanguage)) return replyLanguage;
    if (AI_SIGNATURE_LOCALES.includes(settings.appLanguage)) return settings.appLanguage;
    return "pt-PT";
}

function buildAiSignaturePayload(settings: CockpitSettingsV1, locale: AiLocale): AiSignaturePayload | null {
    const signatureLocale = resolveSignatureLocale(settings, locale);
    const html = String(settings.signaturesHtml?.[signatureLocale] || "").trim();
    const text = String(settings.signatures?.[signatureLocale] || "").trim();
    const storedImage = getSignatureImageDataUrl(signatureLocale);
    const imageUrl = String(storedImage || settings.signatureImageUrl?.[signatureLocale] || "").trim();
    const imageMaxWidth = Math.max(80, Math.min(800, Number(settings.signatureImageMaxWidth?.[signatureLocale] || 260) || 260));

    if (!html && !text && !imageUrl) return null;
    return {
        html: html || undefined,
        text: text || undefined,
        imageUrl: imageUrl || undefined,
        imageMaxWidth,
    };
}

function signaturePayloadToHtml(signature: AiSignaturePayload | null | undefined): string {
    if (!signature) return "";
    const parts: string[] = [];
    const html = String(signature.html || "").trim();
    const text = String(signature.text || "").trim();
    const imageUrl = String(signature.imageUrl || "").trim();
    const imageMaxWidth = Math.max(80, Math.min(800, Number(signature.imageMaxWidth || 260) || 260));

    if (html) {
        parts.push(html);
    } else if (text) {
        parts.push(`<div>${escapeHtml(text).replace(/\n/g, "<br/>")}</div>`);
    }

    if (imageUrl) {
        parts.push(`<div><img src="${escapeHtml(imageUrl)}" alt="" style="max-width:${imageMaxWidth}px;height:auto;border:0;" /></div>`);
    }

    if (!parts.length) return "";
    return `<div data-iccc-signature="official" style="margin-top:16px;">${parts.join("")}</div>`;
}

function appendOfficialSignature(html: string, signature: AiSignaturePayload | null | undefined): string {
    const signatureHtml = signaturePayloadToHtml(signature);
    const source = String(html || "").trim();
    if (!source || !signatureHtml || source.includes('data-iccc-signature="official"')) return source;
    return `${source}${signatureHtml}`;
}

function normalizeForwardSubject(subject: string): string {
    const cleaned = String(subject || "")
        .replace(/^\s*((re|fw|fwd)\s*:\s*)+/i, "")
        .trim();
    return cleaned || String(subject || "").trim();
}

function normalizeReplySubject(subject: string): string {
    const cleaned = String(subject || "").trim();
    if (!cleaned) return "Re:";
    if (/^\s*re\s*:/i.test(cleaned)) return cleaned;
    return `Re: ${normalizeForwardSubject(cleaned) || cleaned}`;
}

function buildTicketEmailSubject(baseSubject: string | undefined, ticketCode: string | undefined, includeTicketCode: boolean): string {
    const code = String(ticketCode || "").trim();
    const raw = String(baseSubject || "").trim();
    if (!includeTicketCode || !code) return raw;
    if (raw.toLowerCase().includes(code.toLowerCase())) return raw;

    const prefixMatch = raw.match(/^((?:(?:re|fw|fwd)\s*:\s*)+)/i);
    const prefixes = prefixMatch?.[1] || "";
    const remainder = raw.slice(prefixes.length).trim();

    if (prefixes) {
        return `${prefixes}[${code}] ${remainder || "Ticket"}`.trim();
    }
    return `[${code}] ${raw || "Ticket"}`.trim();
}

function formatDateLabel(value: string): string {
    const raw = String(value || "").trim();
    if (!raw) return "";
    const parsed = new Date(raw);
    if (Number.isNaN(parsed.getTime())) return raw;
    return parsed.toLocaleString("pt-PT", {
        day: "2-digit",
        month: "2-digit",
        year: "numeric",
        hour: "2-digit",
        minute: "2-digit",
    });
}

async function copyTextWithFallback(text: string): Promise<boolean> {
    const value = String(text || "");
    if (!value) return true;

    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(value);
            return true;
        }
    } catch {
        // Fall back below for Outlook/WebView hosts that block Clipboard API.
    }

    try {
        const ta = document.createElement("textarea");
        ta.value = value;
        ta.setAttribute("readonly", "true");
        ta.style.position = "fixed";
        ta.style.top = "-1000px";
        ta.style.left = "-1000px";
        ta.style.opacity = "0";
        document.body.appendChild(ta);
        ta.focus();
        ta.select();
        ta.setSelectionRange(0, ta.value.length);
        const ok = document.execCommand("copy");
        document.body.removeChild(ta);
        return ok;
    } catch {
        return false;
    }
}

function isSameStoredEmailTarget(
    ctx: any,
    target: AiReplyTargetSelection | null | undefined,
): boolean {
    const currentItemId = String(ctx?.itemId || "").trim();
    const currentMessageId = String(ctx?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
    const targetItemId = String(target?.itemId || "").trim();
    const targetMessageId = String(target?.internetMessageId || "").trim().toLowerCase().replace(/[<>\s]/g, "");
    return Boolean(
        target
        && ((targetItemId && currentItemId && targetItemId === currentItemId) || (targetMessageId && currentMessageId && targetMessageId === currentMessageId))
    );
}

function mergeUniqueStrings(...sources: Array<string[] | undefined>): string[] {
    const seen = new Set<string>();
    const result: string[] = [];
    for (const source of sources) {
        for (const raw of source || []) {
            const value = String(raw || "").trim();
            if (!value) continue;
            const key = value.toLowerCase();
            if (seen.has(key)) continue;
            seen.add(key);
            result.push(value);
        }
    }
    return result;
}

function normalizeExtractedTasks(rawValue: unknown): Array<{ title: string; dueDate?: string; owner?: string }> {
    const list = Array.isArray(rawValue)
        ? rawValue
        : Array.isArray((rawValue as any)?.tasks)
            ? (rawValue as any).tasks
            : [];

    return list
        .map((task: any) => ({
            title: String(task?.title || task?.task || task?.name || "").trim(),
            dueDate: String(task?.dueDate || task?.due || "").trim() || undefined,
            owner: String(task?.owner || task?.assignee || "").trim() || undefined,
        }))
        .filter((task: any) => task.title);
}

function parseExtractedTasks(rawValue: unknown): Array<{ title: string; dueDate?: string; owner?: string }> {
    if (Array.isArray(rawValue) || (rawValue && typeof rawValue === "object")) {
        return normalizeExtractedTasks(rawValue);
    }

    const trimmed = String(rawValue || "").trim();
    if (!trimmed) return [];

    const candidates = [
        trimmed,
        trimmed.replace(/^```(?:json)?\s*/i, "").replace(/\s*```$/i, "").trim(),
    ];

    const arrayStart = trimmed.indexOf("[");
    const arrayEnd = trimmed.lastIndexOf("]");
    if (arrayStart >= 0 && arrayEnd > arrayStart) {
        candidates.push(trimmed.slice(arrayStart, arrayEnd + 1).trim());
    }

    const objectStart = trimmed.indexOf("{");
    const objectEnd = trimmed.lastIndexOf("}");
    if (objectStart >= 0 && objectEnd > objectStart) {
        candidates.push(trimmed.slice(objectStart, objectEnd + 1).trim());
    }

    for (const candidate of candidates) {
        try {
            const parsed = JSON.parse(candidate);
            const normalized = normalizeExtractedTasks(parsed);
            if (normalized.length || Array.isArray(parsed) || Array.isArray(parsed?.tasks)) {
                return normalized;
            }
        } catch {
            // try next candidate
        }
    }

    return [];
}

export const AiCockpit: React.FC = () => {
    const isDevRuntime = window.location.hostname === "localhost" || window.location.hostname === "127.0.0.1";
    const { ctx, bodyText, bodyHtml, links, attachments: liveAttachments, setMsg, aiState, setAiState, files, addFile, removeFile, clearFiles, settings } = useCockpit() as any;
    const aiManualOnly = settings?.aiManualOnly !== false;

    // Local state for immediate typing feel
    // Initialized from context, but NOT synced on every render to avoid loops
    const [prompt, setPrompt] = useState(aiState.prompt);
    const [output, setOutput] = useState(aiState.output);
    const [briefing, setBriefing] = useState<string | null>(null);
    const [isGenerating, setIsGenerating] = useState(false);
    const [isFetchingIntents, setIsFetchingIntents] = useState(false);
    const [isFetchingBriefing, setIsFetchingBriefing] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [isExtractingContacts, setIsExtractingContacts] = useState(false);
    const [briefingExpanded, setBriefingExpanded] = useState(false);
    const [debugLog, setDebugLog] = useState("");
    const [contextBundle, setContextBundle] = useState<AiContextBundle | null>(null);
    const contextBundleCacheRef = useRef<Map<string, AiContextBundle>>(new Map());
    const contextBundlePendingRef = useRef<Map<string, Promise<AiContextBundle | null>>>(new Map());

    // Draft Preview State
    const [draftTo, setDraftTo] = useState<string[]>([]);
    const [draftCc, setDraftCc] = useState<string[]>([]);
    const [draftBcc, setDraftBcc] = useState<string[]>([]);
    const [draftSubject, setDraftSubject] = useState("");
    const [draftTicketCode, setDraftTicketCode] = useState("");
    const [suggestedContacts, setSuggestedContacts] = useState<string[]>([]);
    const [showDraftPreview, setShowDraftPreview] = useState(false);
    const [draftDetailsExpanded, setDraftDetailsExpanded] = useState(false);
    const [extractedTasks, setExtractedTasks] = useState<Array<{ title: string; dueDate?: string; owner?: string; completed?: boolean }>>([]);
    const [showTaskReview, setShowTaskReview] = useState(false);
    const [isExtractingTasks, setIsExtractingTasks] = useState(false);

    // Voice Dictation State
    const [isRecording, setIsRecording] = useState(false);
    const [dictationTarget, setDictationTarget] = useState<"main" | "refine">("main");
    const [refineInput, setRefineInput] = useState("");
    const recognitionRef = useRef<any>(null);

    // Presets Search State
    const [presetSearch, setPresetSearch] = useState("");
    const [intentSearch, setIntentSearch] = useState("");
    const [contactSearch, setContactSearch] = useState("");
    const [fileSearch, setFileSearch] = useState("");

    // History / Rollback
    const [history, setHistory] = useState<HistoryEntry[]>([]);
    const [historyExpanded, setHistoryExpanded] = useState(false);
    const emailKey = getEmailKey(ctx);
    const [activePanel, setActivePanel] = useState<QuickPanelId>(null);
    const [selectedCustomToneId, setSelectedCustomToneId] = useState<string>("");
    const [replyTargetEmail, setReplyTargetEmail] = useState<AiReplyTargetSelection | null>(null);
    const [replyAddresseeName, setReplyAddresseeName] = useState("");
    const [replyAddresseeContext, setReplyAddresseeContext] = useState("");
    const [generationError, setGenerationError] = useState("");
    const [fileUsage, setFileUsage] = useState<Record<string, FileUsageState>>({});
    const [persistedCurrentEmail, setPersistedCurrentEmail] = useState<RelatedEmailEntry | null>(null);

    // Responsive UI Scale
    const paneRef = useRef<HTMLDivElement>(null);
    const [uiScale, setUiScale] = useState(1);
    const [paneWidth, setPaneWidth] = useState(0);

    useEffect(() => {
        if (!paneRef.current) return;
        const observer = new ResizeObserver((entries) => {
            for (const entry of entries) {
                const width = entry.contentRect.width;
                setPaneWidth(width);
                if (width < 320) setUiScale(0.82);
                else if (width < 360) setUiScale(0.88);
                else if (width < 400) setUiScale(0.94);
                else setUiScale(1);
            }
        });
        observer.observe(paneRef.current);
        return () => observer.disconnect();
    }, []);

    const isNarrow = paneWidth < 360;
    const px = (n: number) => `${Math.max(20, Math.round(n * uiScale))}px`;
    const fpx = (n: number) => `${Math.max(9, Math.round(n * uiScale))}px`;

    // CRITICAL: Only sync local state from context when the conversation (email) changes.
    useEffect(() => {
        setPrompt(aiState.prompt);
        setOutput(aiState.output);
        setDebugLog(""); // Clear debug log on switch
        const nextHistory = loadHistory().filter(h => h.emailKey === emailKey);
        setHistory(nextHistory);
    }, [ctx.conversationId, emailKey]);

    useEffect(() => {
        setBriefing(null);
        setBriefingExpanded(false);
        setSuggestedContacts([]);
        setExtractedTasks([]);
        setShowTaskReview(false);
        setContactSearch("");
        setIntentSearch("");
        setFileSearch("");
        setContextBundle(null);
        setActivePanel(null);
        setHistoryExpanded(false);
        setSelectedCustomToneId("");
        setReplyTargetEmail(null);
        setReplyAddresseeName("");
        setReplyAddresseeContext("");
        setDraftTo([]);
        setDraftCc([]);
        setDraftBcc([]);
        setDraftSubject("");
        setDraftTicketCode("");
        setShowDraftPreview(false);
        setDraftDetailsExpanded(false);
        setFileUsage({});
    }, [emailKey]);

    const currentEmailBootstrapPayload = useMemo<RelevantEmailPayload>(() => ({
        itemId: String(ctx?.itemId || "").trim(),
        internetMessageId: String(ctx?.internetMessageId || "").trim(),
        conversationId: String(ctx?.conversationId || "").trim(),
        subject: String(ctx?.subject || "").trim(),
        fromEmail: String(ctx?.fromEmail || "").trim(),
        fromName: String(ctx?.fromName || "").trim(),
        receivedAtIso: String(ctx?.receivedDateTimeIso || "").trim(),
        messageDateIso: String(ctx?.receivedDateTimeIso || "").trim(),
        bodyText: String(bodyText || "").trim(),
        bodyHtml: String(bodyHtml || "").trim(),
        attachments: (liveAttachments || []).map((attachment: any) => ({
            key: attachment?.key,
            id: attachment?.id,
            name: String(attachment?.name || "").trim(),
            contentType: String(attachment?.contentType || "application/octet-stream").trim(),
            size: Number(attachment?.size || 0) || undefined,
            isInline: Boolean(attachment?.isInline),
            contentId: String(attachment?.contentId || "").trim() || undefined,
            content: String(attachment?.content || "").trim(),
            hasContent: Boolean(String(attachment?.content || "").trim()),
        })).filter((attachment: any) => attachment.name),
    }), [bodyHtml, bodyText, ctx?.conversationId, ctx?.fromEmail, ctx?.fromName, ctx?.internetMessageId, ctx?.itemId, ctx?.receivedDateTimeIso, ctx?.subject, liveAttachments]);

    const currentEmailBootstrapLinkPayload = useMemo<RelevantEmailPayload>(() => ({
        ...currentEmailBootstrapPayload,
        attachments: (currentEmailBootstrapPayload.attachments || []).map(({ content: _content, ...attachment }) => attachment),
    }), [currentEmailBootstrapPayload]);

    const persistedEmailAttachments = useMemo<PersistedEmailAttachment[]>(
        () => Array.isArray(persistedCurrentEmail?.attachments)
            ? persistedCurrentEmail.attachments.filter((attachment) => String(attachment?.name || "").trim() && !isRejectedAttachmentState(attachment?.documentState))
            : [],
        [persistedCurrentEmail]
    );

    const effectiveBodyText = String(persistedCurrentEmail?.bodyText || bodyText || "").trim();
    const effectiveBodyHtml = String(persistedCurrentEmail?.bodyHtml || bodyHtml || "").trim();

    useEffect(() => {
        setPersistedCurrentEmail(null);
    }, [emailKey]);

    useEffect(() => {
        let cancelled = false;
        const hasIdentity = Boolean(
            currentEmailBootstrapLinkPayload.itemId
            || currentEmailBootstrapLinkPayload.internetMessageId
            || currentEmailBootstrapLinkPayload.conversationId
            || currentEmailBootstrapLinkPayload.subject
        );
        if (!hasIdentity) {
            setPersistedCurrentEmail(null);
            return;
        }

        const loadPersistedEmail = async () => {
            const loadRelated = async () =>
                getRelatedEmailContext(currentEmailBootstrapLinkPayload).catch(() => null as Awaited<ReturnType<typeof getRelatedEmailContext>> | null);

            const pickBestEmail = (context: Awaited<ReturnType<typeof getRelatedEmailContext>> | null): RelatedEmailEntry | null => {
                const rows = [
                    context?.email,
                    ...((context?.emails || []).filter(Boolean) as RelatedEmailEntry[]),
                ].filter(Boolean) as RelatedEmailEntry[];
                return rows.find((email) => isSameStoredEmailTarget(ctx, email as any)) || rows[0] || null;
            };

            let related = await loadRelated();
            let email = pickBestEmail(related);
            const needsRegistration = !email || (
                Array.isArray(currentEmailBootstrapPayload.attachments)
                && currentEmailBootstrapPayload.attachments.length > 0
                && !(Array.isArray(email?.attachments) && email.attachments.length > 0)
            );

            if (needsRegistration) {
                await registerRelevantEmail({
                    ...currentEmailBootstrapPayload,
                    ...getGroupAttachmentStorageOptions(settings),
                }).catch(() => null);
                related = await loadRelated();
                email = pickBestEmail(related);
            }

            if (!cancelled) {
                setPersistedCurrentEmail(email || null);
            }
        };

        void loadPersistedEmail();

        return () => {
            cancelled = true;
        };
    }, [
        currentEmailBootstrapLinkPayload.conversationId,
        currentEmailBootstrapLinkPayload.internetMessageId,
        currentEmailBootstrapLinkPayload.itemId,
        currentEmailBootstrapLinkPayload.subject,
        currentEmailBootstrapPayload,
        ctx,
        settings?.groupStorage?.baseFolderPath,
        settings?.groupStorage?.mode,
        settings?.groupStorage?.provider,
    ]);

    useEffect(() => {
        const nextUsage: Record<string, FileUsageState> = {};
        for (const file of files || []) {
            const name = String(file?.name || "").trim();
            if (!name) continue;
            nextUsage[name] = { analyze: true, forward: false };
        }
        setFileUsage((prev) => {
            const merged: Record<string, FileUsageState> = {};
            for (const att of persistedEmailAttachments || []) {
                const name = String(att?.name || "").trim();
                if (!name) continue;
                merged[name] = prev[name] || nextUsage[name] || { analyze: false, forward: false };
            }
            for (const [name, flags] of Object.entries(nextUsage)) {
                merged[name] = merged[name] || flags;
            }
            return merged;
        });
    }, [files, persistedEmailAttachments]);

    const selectedAction: AiAction = aiState.action === "forward" ? "forward" : "reply";

    useEffect(() => {
        if (!Array.isArray(persistedEmailAttachments) || persistedEmailAttachments.length === 0) return;
        const shouldPrimeForward = selectedAction === "forward" && !Object.values(fileUsage || {}).some((flags) => flags?.forward);
        let changed = false;

        setFileUsage((prev) => {
            const next: Record<string, FileUsageState> = { ...(prev || {}) };
            for (const attachment of persistedEmailAttachments) {
                const name = String(attachment?.name || "").trim();
                if (!name) continue;

                if (!next[name]) {
                    next[name] = { analyze: true, forward: shouldPrimeForward };
                    syncAttachmentSelection(name, next[name]);
                    changed = true;
                    continue;
                }

                if (shouldPrimeForward && !next[name].forward) {
                    next[name] = { ...next[name], analyze: true, forward: true };
                    syncAttachmentSelection(name, next[name]);
                    changed = true;
                }
            }
            return changed ? next : prev;
        });
    }, [persistedEmailAttachments, selectedAction]);

    function hasEmailIdentity() {
        return Boolean(ctx.itemId || ctx.internetMessageId || ctx.conversationId);
    }

    async function resolvePersistedAttachmentContent(attachment: PersistedEmailAttachment | null | undefined): Promise<string> {
        const localContent = String(attachment?.content || "").trim();
        if (localContent) return localContent;
        const emailId = String(persistedCurrentEmail?.id || "").trim();
        const attachmentId = String(attachment?.key || attachment?.id || "").trim();
        if (!emailId || !attachmentId || attachment?.hasContent !== true) return "";
        try {
            const loaded = await getEmailAttachmentContentBase64(emailId, attachmentId);
            const content = String(loaded?.base64 || "").trim();
            if (content) {
                setPersistedCurrentEmail((current) => {
                    if (!current || String(current.id || "").trim() !== emailId || !Array.isArray(current.attachments)) return current;
                    return {
                        ...current,
                        attachments: current.attachments.map((entry) => {
                            const entryId = String(entry?.key || entry?.id || "").trim();
                            if (entryId !== attachmentId) return entry;
                            return {
                                ...entry,
                                content,
                                hasContent: true,
                            };
                        }),
                    };
                });
            }
            return content;
        } catch {
            return "";
        }
    }

    async function resolveSelectedAnalyzeFiles(): Promise<Array<{ name: string; type: string; content: string }>> {
        const selectedNames = new Set(
            Object.entries(fileUsage || {})
                .filter(([, flags]) => flags?.analyze)
                .map(([name]) => String(name || "").trim())
                .filter(Boolean)
        );
        const results: Array<{ name: string; type: string; content: string }> = [];
        const handled = new Set<string>();

        for (const attachment of persistedEmailAttachments) {
            const name = String(attachment?.name || "").trim();
            if (!name || !selectedNames.has(name)) continue;
            const content = await resolvePersistedAttachmentContent(attachment);
            if (!content) continue;
            results.push({
                name,
                type: String(attachment?.contentType || "application/octet-stream").trim() || "application/octet-stream",
                content,
            });
            handled.add(name);
        }

        for (const file of files || []) {
            const name = String(file?.name || "").trim();
            const content = String(file?.content || "").trim();
            if (!name || !content || handled.has(name) || !selectedNames.has(name)) continue;
            results.push({
                name,
                type: String(file?.type || "application/octet-stream").trim() || "application/octet-stream",
                content,
            });
        }

        return results;
    }

    async function resolveSelectedForwardFiles(): Promise<Array<{ name: string; type: string; content: string }>> {
        const selectedNames = new Set(
            Object.entries(fileUsage || {})
                .filter(([, flags]) => flags?.forward)
                .map(([name]) => String(name || "").trim())
                .filter(Boolean)
        );
        const results: Array<{ name: string; type: string; content: string }> = [];
        for (const attachment of persistedEmailAttachments) {
            const name = String(attachment?.name || "").trim();
            if (!name || !selectedNames.has(name)) continue;
            const content = await resolvePersistedAttachmentContent(attachment);
            if (!content) continue;
            results.push({
                name,
                type: String(attachment?.contentType || "application/octet-stream").trim() || "application/octet-stream",
                content,
            });
        }
        return results;
    }

    async function ensureContextBundle(force = false): Promise<AiContextBundle | null> {
        if (!hasEmailIdentity()) return null;

        if (!force) {
            const cached = contextBundleCacheRef.current.get(emailKey);
            if (cached) {
                if (!contextBundle) setContextBundle(cached);
                return cached;
            }
            const pending = contextBundlePendingRef.current.get(emailKey);
            if (pending) return pending;
        }

        const promise = buildAiContextBundle({
            ctx,
            bodyText: effectiveBodyText,
            bodyHtml: effectiveBodyHtml,
            links: Array.isArray(links) ? links : [],
            attachments: Array.isArray(persistedEmailAttachments) ? persistedEmailAttachments as any : [],
        })
            .then((bundle) => {
                contextBundleCacheRef.current.set(emailKey, bundle);
                setContextBundle(bundle);
                return bundle;
            })
            .catch((error) => {
                if (force) {
                    console.error("[AiCockpit] Failed to build AI context bundle:", error);
                }
                return null;
            })
            .finally(() => {
                contextBundlePendingRef.current.delete(emailKey);
            });

        contextBundlePendingRef.current.set(emailKey, promise);
        return promise;
    }

    useEffect(() => {
        if (aiManualOnly || !ctx.conversationId || ctx.isCompose) return;

        void handleFetchBriefing(false);
        void handleFetchIntents(false);
        if (effectiveBodyText) void handleExtractContacts(false);
        if ((effectiveBodyText || "").length >= 50) void handleExtractTasksReview();
    }, [aiManualOnly, emailKey, ctx.conversationId, ctx.isCompose, effectiveBodyText]);

    useEffect(() => {
        if (!settings) return;
        if (!aiState.prompt && !aiState.output && !aiState.history.length) {
            setAiState({
                tone: settings.tone || "neutro",
                locale: (settings.replyLanguage || "auto") as AiLocale,
                action: "reply",
            });
        }
    }, [emailKey, settings]);

    const selectedLocale = ((aiState.locale || settings?.replyLanguage || "auto") as AiLocale);
    const effectiveLocale = (selectedLocale !== "auto"
        ? selectedLocale
        : ((settings?.readingLanguage && settings.readingLanguage !== "auto"
            ? settings.readingLanguage
            : (settings?.appLanguage || "pt-PT")) as AiLocale));

    const baseTone = aiState.tone || settings?.tone || "neutro";
    const currentCustomTone = (settings?.aiCustomTones || []).find((entry: any) => entry.id === selectedCustomToneId) || null;
    const selectedAnalyzeFiles = (files || [])
        .filter((entry: any) => fileUsage[String(entry?.name || "").trim()]?.analyze)
        .map((entry: any) => ({
            name: entry.name,
            type: entry.type || entry.contentType,
            content: entry.content || "",
        }))
        .filter((entry: { name?: string; content?: string }) => entry.name && entry.content);
    const selectedForwardFiles = (persistedEmailAttachments || [])
        .filter((entry) => fileUsage[String(entry?.name || "").trim()]?.forward)
        .map((entry) => ({
            name: entry.name,
            type: entry.contentType,
            content: entry.content || "",
            hasContent: entry.hasContent,
        }))
        .filter((entry) => entry.name && (entry.content || entry.hasContent));
    const availableAttachmentCount = (persistedEmailAttachments || [])
        .filter((entry) => String(entry?.name || "").trim() && (String(entry?.content || "").trim() || entry?.hasContent))
        .length;

    function setGenerationAction(nextAction: "reply" | "forward") {
        setAiState({ action: nextAction });
    }

    function applyHistoryEntry(entry: HistoryEntry) {
        setOutput(entry.output || "");
        setPrompt(entry.prompt || "");
        setDraftTo(normalizeEmailListInput(entry.draftTo));
        setDraftCc(normalizeEmailListInput(entry.draftCc));
        setDraftBcc(normalizeEmailListInput(entry.draftBcc));
        setDraftSubject(String(entry.draftSubject || ""));
        setSelectedCustomToneId(String(entry.customToneId || ""));
        setReplyTargetEmail(entry.replyTarget || null);
        setReplyAddresseeName(String(entry.replyDirection?.addresseeName || ""));
        setReplyAddresseeContext(String(entry.replyDirection?.addresseeContext || ""));
        setShowDraftPreview(true);
        setDraftDetailsExpanded(true);
        setAiState({
            prompt: entry.prompt || "",
            output: entry.output || "",
            action: entry.action || "reply",
            tone: entry.tone || "neutro",
            locale: entry.locale || "auto",
            suggestedTo: normalizeEmailListInput(entry.draftTo),
            suggestedCc: normalizeEmailListInput(entry.draftCc),
            suggestedSubject: String(entry.draftSubject || ""),
        });
    }

    function syncAttachmentSelection(name: string, nextFlags: FileUsageState) {
        const fileName = String(name || "").trim();
        if (!fileName) return;
        const sourceAttachment = (persistedEmailAttachments || []).find((entry) => String(entry?.name || "").trim() === fileName);
        if (!sourceAttachment) return;

        const shouldKeep = Boolean(nextFlags.analyze || nextFlags.forward);
        const exists = (files || []).some((entry: any) => String(entry?.name || "").trim() === fileName);
        if (shouldKeep && !exists) {
            const immediateContent = String(sourceAttachment.content || "").trim();
            if (immediateContent) {
                addFile({
                    name: sourceAttachment.name,
                    type: sourceAttachment.contentType,
                    content: immediateContent,
                });
            } else {
                void resolvePersistedAttachmentContent(sourceAttachment).then((content) => {
                    if (!content) return;
                    addFile({
                        name: sourceAttachment.name,
                        type: sourceAttachment.contentType,
                        content,
                    });
                });
            }
        }
        if (!shouldKeep && exists) {
            removeFile(fileName);
        }
    }

    function setAttachmentUsage(name: string, patch: Partial<FileUsageState>) {
        const fileName = String(name || "").trim();
        if (!fileName) return;
        setFileUsage((prev) => {
            const current = prev[fileName] || { analyze: false, forward: false };
            const next = { ...current, ...patch };
            syncAttachmentSelection(fileName, next);
            return { ...prev, [fileName]: next };
        });
    }

    async function handleExtractTasksReview() {
        const effectiveBodyTextForTasks = effectiveBodyText || htmlToPlainText(effectiveBodyHtml || "");
        if (!ctx.conversationId || ctx.isCompose || isExtractingTasks) return;
        if (!effectiveBodyTextForTasks) {
            setMsg("O corpo deste email ainda não ficou disponível no Outlook. Tenta novamente dentro de 1-2 segundos.");
            return;
        }
        if (effectiveBodyTextForTasks.length < 50) {
            setMsg("O email e demasiado curto para detetar tarefas com confianca.");
            return;
        }

        setIsExtractingTasks(true);
        try {
            const res = await aiGenerate({
                action: "extract_tasks_json" as any,
                mode: "fast",
                locale: (aiState.locale || "auto") as any,
                tone: "neutro",
                email: {
                    subject: ctx.subject || "",
                    from: ctx.fromEmail || "",
                    to: (ctx.toRecipients || []).map((r: any) => r.email),
                    cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                    bodyText: effectiveBodyTextForTasks,
                } as any
            });
            if (!res.ok) {
                setExtractedTasks([]);
                setShowTaskReview(false);
                setMsg(res.error || "Erro ao extrair tarefas.");
                return;
            }

            const tasks = parseExtractedTasks((res as any).data ?? res.text ?? "");
            if (tasks.length > 0) {
                setExtractedTasks(tasks.map((t) => ({ ...t, completed: false })));
                setShowTaskReview(true);
            } else {
                console.error("Failed to parse tasks JSON:", res.text);
                setExtractedTasks([]);
                setShowTaskReview(false);
                setMsg("Nao foram encontradas tarefas concretas neste email.");
            }
        } catch (err) {
            console.error("Erro ao extrair tarefas:", err);
            setMsg("Erro ao extrair tarefas.");
        } finally {
            setIsExtractingTasks(false);
        }
    }

    async function handleFetchIntents(force = false) {
        if (!ctx.conversationId || ctx.isCompose || isFetchingIntents) return;
        if (!force && aiState.smartReplies.length > 0) return;

        setIsFetchingIntents(true);
        try {
            const nextSettings = await getSettings();
            const bundle = await ensureContextBundle(false);
            const effectiveBodyTextForIntents = effectiveBodyText || htmlToPlainText(effectiveBodyHtml || "");
            const res = await aiGenerate({
                action: "intent_proposals",
                mode: "fast",
                locale: (nextSettings.replyLanguage || "pt-PT") as any,
                tone: nextSettings.tone || "neutro",
                email: {
                    subject: ctx.subject || "",
                    from: ctx.fromEmail || "",
                    to: (ctx.toRecipients || []).map((r: any) => r.email),
                    cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                    bodyText: effectiveBodyTextForIntents,
                    bodyScope: nextSettings.bodyScope || "main"
                },
                briefing: briefing,
                contextBundle: bundle?.promptContext || "",
                persona: {
                    userRole: nextSettings.userRole,
                    styleContext: nextSettings.styleContext,
                    styleExamples: nextSettings.styleExamples,
                }
            });
            if (res.ok) {
                const intents = res.text.split(";").map((item) => item.trim()).filter(Boolean);
                setAiState({ smartReplies: intents });
                if (!intents.length) setMsg("Nao surgiram sugestoes rapidas para este email.");
            }
        } catch (err) {
            console.error("Erro ao obter intencoes:", err);
            setMsg("Erro ao obter sugestoes rapidas.");
        } finally {
            setIsFetchingIntents(false);
        }
    }

    async function handleFetchBriefing(force = false) {
        if (!ctx.conversationId || ctx.isCompose || isFetchingBriefing) return;
        if (!force && briefing) return;

        try {
            setIsFetchingBriefing(true);
            const { aiGenerateBriefing } = await import("@/api");
            const bundle = await ensureContextBundle(true);
            const effectiveBodyTextForBriefing = effectiveBodyText || htmlToPlainText(effectiveBodyHtml || "");
            const briefingContext = bundle?.briefingContext || effectiveBodyTextForBriefing;
            const briefingCacheKey = bundle?.cacheKey
                ? `${ctx.conversationId || emailKey}|${bundle.cacheKey}`
                : (ctx.conversationId || emailKey);
            const res = await aiGenerateBriefing(briefingContext, [], {}, ctx.conversationId, briefingCacheKey);
            if (res.ok) {
                setBriefing(res.summary || "");
            }
        } catch (err) {
            console.error("Erro ao obter briefing:", err);
            setMsg("Erro ao gerar briefing.");
        } finally {
            setIsFetchingBriefing(false);
        }
    }

    async function handleExtractContacts(force = false) {
        if (!bodyText || ctx.isCompose || isExtractingContacts) return;
        if (!force && suggestedContacts.length > 0) return;

        setIsExtractingContacts(true);
        try {
            const res = await aiGenerate({
                action: "extract_contacts" as any,
                mode: "fast",
                locale: (aiState.locale || "auto") as any,
                tone: "neutro",
                email: {
                    bodyText: effectiveBodyText,
                    subject: "",
                    from: "",
                    to: [],
                    cc: []
                } as any
            });
            if (res.ok && res.text) {
                const emails = res.text.split(";").map((item) => item.trim()).filter(Boolean);
                setSuggestedContacts(emails);
                if (!emails.length) setMsg("Nao foram detetados contactos adicionais neste email.");
            }
        } catch (err) {
            console.error("Erro ao extrair contactos:", err);
            setMsg("Erro ao extrair contactos.");
        } finally {
            setIsExtractingContacts(false);
        }
    }

    function toggleIntentsMenu() {
        const nextOpen = activePanel !== "intents";
        setActivePanel(nextOpen ? "intents" : null);
        if (nextOpen && !aiState.smartReplies.length) {
            void handleFetchIntents(true);
        }
    }

    function toggleContactsMenu() {
        const nextOpen = activePanel !== "contacts";
        setActivePanel(nextOpen ? "contacts" : null);
        if (nextOpen && !suggestedContacts.length) {
            void handleExtractContacts(true);
        }
    }

    async function handlePickReplyTarget() {
        try {
            const selection = await openAiReplyTargetPicker({
                conversationId: String(ctx.conversationId || ""),
                internetMessageId: String(ctx.internetMessageId || ""),
                itemId: String(ctx.itemId || ""),
                subject: String(ctx.subject || ""),
                fromEmail: String(ctx.fromEmail || ""),
                fromName: String(ctx.fromName || ""),
                receivedAtIso: String(ctx.receivedDateTimeIso || ""),
                selectedEmailKey: String(replyTargetEmail?.emailKey || ""),
            });
            if (!selection?.emailKey) return;
            setReplyTargetEmail(selection);
            setShowDraftPreview(true);
            setDraftDetailsExpanded(true);
            setMsg(`Email-alvo selecionado: ${selection.subject || selection.fromEmail || selection.emailKey}`);
        } catch (error: any) {
            setMsg(String(error?.message || error || "Nao foi possivel abrir a selecao de emails relacionados."));
        }
    }

    function buildPayloadFromCurrent(): RelevantEmailPayload {
        return {
            itemId: String(ctx.itemId || "").trim() || undefined,
            internetMessageId: String(ctx.internetMessageId || "").trim() || undefined,
            conversationId: String(ctx.conversationId || "").trim() || undefined,
            subject: String(ctx.subject || "").trim() || undefined,
            fromEmail: String(ctx.fromEmail || "").trim() || undefined,
            fromName: String(ctx.fromName || "").trim() || undefined,
            receivedAtIso: String(ctx.receivedDateTimeIso || "").trim() || undefined,
            messageDateIso: String(ctx.receivedDateTimeIso || "").trim() || undefined,
        };
    }

    function buildPayloadFromTarget(): RelevantEmailPayload | null {
        if (!replyTargetEmail) return null;
        return {
            itemId: String(replyTargetEmail.itemId || "").trim() || undefined,
            internetMessageId: String(replyTargetEmail.internetMessageId || "").trim() || undefined,
            conversationId: String(replyTargetEmail.conversationId || "").trim() || undefined,
            subject: String(replyTargetEmail.subject || "").trim() || undefined,
            fromEmail: String(replyTargetEmail.fromEmail || "").trim() || undefined,
            fromName: String(replyTargetEmail.fromName || "").trim() || undefined,
            receivedAtIso: String(replyTargetEmail.receivedAtIso || replyTargetEmail.messageDateIso || "").trim() || undefined,
            messageDateIso: String(replyTargetEmail.messageDateIso || replyTargetEmail.receivedAtIso || "").trim() || undefined,
        };
    }

    async function collectDraftCategorySource(payload: RelevantEmailPayload | null, fallbackHasOdooLinks = false): Promise<OutlookCategorySource> {
        if (!payload) {
            return buildOutlookCategorySourceFromRelatedContext({
                email: null,
                groups: [],
                tickets: [],
                settings,
                specialCategories: fallbackHasOdooLinks ? [ODOO_LINKED_CATEGORY] : [],
                managedSpecialCategories: [ODOO_LINKED_CATEGORY],
            });
        }
        const response = await getRelatedEmailContext(payload);
        const hasOdooLinks = fallbackHasOdooLinks || Boolean(response?.email?.relatedRecords?.length);
        return buildOutlookCategorySourceFromRelatedContext({
            email: response?.email || null,
            groups: Array.isArray(response?.groups) ? response.groups : [],
            tickets: Array.isArray(response?.tickets) ? response.tickets : [],
            settings,
            specialCategories: hasOdooLinks ? [ODOO_LINKED_CATEGORY] : [],
            managedSpecialCategories: [ODOO_LINKED_CATEGORY],
        });
    }

    function pickPreferredTicketCode(currentSnapshot: { ticketCodes: string[] }, targetSnapshot: { ticketCodes: string[] }, shouldIncludeTarget: boolean): string {
        const currentCodes = mergeUniqueStrings(currentSnapshot.ticketCodes);
        const targetCodes = shouldIncludeTarget ? mergeUniqueStrings(targetSnapshot.ticketCodes) : [];
        const sharedCodes = currentCodes.filter((code) => targetCodes.some((targetCode) => targetCode.toLowerCase() === code.toLowerCase()));
        if (sharedCodes.length === 1) return sharedCodes[0];
        if (targetCodes.length === 1) return targetCodes[0];
        if (currentCodes.length === 1) return currentCodes[0];
        const mergedCodes = mergeUniqueStrings(currentCodes, targetCodes);
        return mergedCodes.length === 1 ? mergedCodes[0] : "";
    }

    async function loadDraftLinkMetadata(): Promise<{
        source: OutlookCategorySource;
        preferredTicketCode: string;
    }> {
        const currentSnapshot = await collectDraftCategorySource(buildPayloadFromCurrent(), links.length > 0);
        const shouldIncludeTarget = Boolean(replyTargetEmail && !isSameStoredEmailTarget(ctx, replyTargetEmail));
        const targetSnapshot = shouldIncludeTarget
            ? await collectDraftCategorySource(buildPayloadFromTarget(), false)
            : {
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
            };

        return {
            source: mergeOutlookCategorySources(currentSnapshot, shouldIncludeTarget ? targetSnapshot : null),
            preferredTicketCode: pickPreferredTicketCode(currentSnapshot, targetSnapshot, shouldIncludeTarget),
        };
    }

    async function loadDraftLinkCategories(): Promise<OutlookCategorySource | null> {
        const metadata = await loadDraftLinkMetadata();
        const source = metadata.source;
        const hasAnyCategories = Boolean(
            source.principalGroupNames.length
            || source.referenceGroupNames.length
            || source.ticketCodes.length
            || source.labelNames.length
            || source.groupStatuses.length
            || source.ticketStatuses.length
            || source.labelStatuses.length
            || source.specialCategories.length
            || source.managedLabelNames.length
            || source.managedSpecialCategories.length
        );
        return hasAnyCategories ? source : null;
    }

    useEffect(() => {
        if (settings?.groupTicketUi?.includeTicketCodeInSubject === false) {
            setDraftTicketCode("");
            return;
        }

        let cancelled = false;
        loadDraftLinkMetadata()
            .then((metadata) => {
                if (cancelled) return;
                setDraftTicketCode(String(metadata.preferredTicketCode || "").trim());
            })
            .catch(() => {
                if (!cancelled) setDraftTicketCode("");
            });

        return () => {
            cancelled = true;
        };
    }, [
        ctx.conversationId,
        ctx.fromEmail,
        ctx.fromName,
        ctx.internetMessageId,
        ctx.itemId,
        ctx.receivedDateTimeIso,
        ctx.subject,
        links.length,
        replyTargetEmail?.conversationId,
        replyTargetEmail?.internetMessageId,
        replyTargetEmail?.itemId,
        replyTargetEmail?.subject,
        settings?.groupTicketUi?.includeTicketCodeInSubject,
    ]);

    // Sync draft defaults from context OR persistent aiState
    useEffect(() => {
        const hasSuggestedRecipients = Array.isArray(aiState.suggestedTo) && aiState.suggestedTo.length > 0;
        const hasSuggestedCc = Array.isArray(aiState.suggestedCc) && aiState.suggestedCc.length > 0;
        const hasSuggestedSubject = Boolean(String(aiState.suggestedSubject || "").trim());
        const includeTicketCodeInSubject = settings?.groupTicketUi?.includeTicketCodeInSubject !== false;
        const applyDraftSubjectTicketCode = (subject: string) => buildTicketEmailSubject(subject, draftTicketCode, includeTicketCodeInSubject);

        if (selectedAction === "forward") {
            setDraftSubject(
                applyDraftSubjectTicketCode(
                    String(aiState.suggestedSubject || "").trim() || normalizeForwardSubject(ctx.subject || replyTargetEmail?.subject || "")
                )
            );
            return;
        }

        if (replyTargetEmail?.fromEmail) {
            if (hasSuggestedRecipients || hasSuggestedCc || hasSuggestedSubject) {
                setDraftTo(normalizeEmailListInput(aiState.suggestedTo?.length ? aiState.suggestedTo : [replyTargetEmail.fromEmail]));
                setDraftCc(normalizeEmailListInput(aiState.suggestedCc));
                setDraftSubject(
                    applyDraftSubjectTicketCode(
                        String(aiState.suggestedSubject || "").trim() || normalizeReplySubject(replyTargetEmail.subject || ctx.subject || "")
                    )
                );
            } else {
                setDraftTo(normalizeEmailListInput([replyTargetEmail.fromEmail]));
                setDraftCc([]);
                setDraftSubject(applyDraftSubjectTicketCode(normalizeReplySubject(replyTargetEmail.subject || ctx.subject || "")));
            }
            return;
        }

        if (hasSuggestedRecipients || hasSuggestedCc || hasSuggestedSubject) {
            setDraftTo(normalizeEmailListInput(aiState.suggestedTo));
            setDraftCc(normalizeEmailListInput(aiState.suggestedCc));
            setDraftSubject(applyDraftSubjectTicketCode(aiState.suggestedSubject || ""));
        } else {
            setDraftTo(normalizeEmailListInput([ctx.fromEmail]));
            setDraftCc(normalizeEmailListInput((ctx.ccRecipients || []).map((r: any) => r.email)));
            setDraftSubject(applyDraftSubjectTicketCode(ctx.subject || ""));
        }
    }, [ctx, aiState.suggestedSubject, aiState.suggestedTo, aiState.suggestedCc, selectedAction, replyTargetEmail, draftTicketCode, settings?.groupTicketUi?.includeTicketCodeInSubject]);

    const handlePromptChange = (val: string) => {
        setPrompt(val);
        setAiState({ prompt: val });
    };

    // --- Voice Dictation Logic ---
    const toggleRecording = () => {
        if (isRecording) {
            recognitionRef.current?.stop();
            return;
        }

        const SpeechRecognition = (window as any).SpeechRecognition || (window as any).webkitSpeechRecognition;
        if (!SpeechRecognition) {
            setMsg("O teu browser não suporta reconhecimento de voz.");
            return;
        }

        const recognition = new SpeechRecognition();
        recognition.lang = selectedLocale === "auto" ? (effectiveLocale || "pt-PT") as any : selectedLocale as any;
        recognition.continuous = true;
        recognition.interimResults = true;

        recognition.onstart = () => {
            setIsRecording(true);
            setMsg("A ouvir... (fala agora)");
        };

        recognition.onend = () => {
            setIsRecording(false);
            setMsg("");
        };

        recognition.onerror = (event: any) => {
            console.error("Speech error", event.error);
            setIsRecording(false);
            setMsg("Erro no reconhecimento de voz.");
        };

        recognition.onresult = (event: any) => {
            let newFinal = "";
            for (let i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    newFinal += event.results[i][0].transcript;
                }
            }
            if (newFinal) {
                if (dictationTarget === "main") {
                    setPrompt((prev: string) => {
                        const updated = (prev + " " + newFinal).trim();
                        setAiState({ prompt: updated });
                        return updated;
                    });
                } else {
                    setRefineInput((prev: string) => (prev + " " + newFinal).trim());
                }
            }
        };

        recognitionRef.current = recognition;
        recognition.start();
    };

    const handleImportAttachments = async () => {
        try {
            setIsImporting(true);
            const ctxFiles = files || [];
            const persisted = await getRelatedEmailContext({
                itemId: String(ctx?.itemId || "").trim(),
                internetMessageId: String(ctx?.internetMessageId || "").trim(),
                conversationId: String(ctx?.conversationId || "").trim(),
                subject: String(ctx?.subject || "").trim(),
                fromEmail: String(ctx?.fromEmail || "").trim(),
                receivedAtIso: String(ctx?.receivedDateTimeIso || "").trim(),
            }).catch(() => null);
            const persistedEmail = persisted?.email || null;
            const persistedAttachments = Array.isArray(persistedEmail?.attachments)
                ? persistedEmail.attachments.filter((attachment: any) => !isRejectedAttachmentState(attachment?.documentState))
                : [];
            let persistedCount = 0;
            if (persistedEmail?.id && persistedAttachments.length) {
                for (const attachment of persistedAttachments) {
                    const name = String(attachment?.name || "").trim();
                    if (!name || ctxFiles.find((f: any) => f.name === name)) continue;
                    let content = String(attachment?.content || "").trim();
                    if (!content && attachment?.hasContent) {
                        const remoteId = String(attachment?.key || attachment?.id || "").trim();
                        if (remoteId) {
                            try {
                                const loaded = await getEmailAttachmentContentBase64(String(persistedEmail.id || "").trim(), remoteId);
                                content = String(loaded.base64 || "").trim();
                            } catch {
                                content = "";
                            }
                        }
                    }
                    if (!content) continue;
                    addFile({
                        name,
                        type: String(attachment?.contentType || "application/octet-stream").trim(),
                        content
                    });
                    persistedCount++;
                }
            }
            if (persistedCount > 0) {
                setMsg(`${persistedCount} anexos importados!`);
                return;
            }
            if (persistedEmail?.id) {
                setMsg("Nenhum anexo persistido disponivel para importar.");
            } else {
                setMsg("O email atual ainda nao esta pronto no servidor. Reabre o email e tenta novamente.");
            }
        } catch (e: any) {
            setMsg("Erro ao importar: " + e.message);
        } finally {
            setIsImporting(false);
        }
    };

    const handleAddToKnowledge = async (text: string) => {
        if (!text || !settings) return;
        try {
            const { saveSettings } = await import("@/settings");
            const newKnowledge = [...(settings.aiKnowledge || []), text.trim()];
            await saveSettings({ aiKnowledge: newKnowledge });
            setMsg("Guardado na Base de Conhecimento!");
        } catch (err) {
            console.error("Erro ao guardar conhecimento:", err);
            setMsg("Erro ao guardar.");
        }
    };

    async function handleGenerate(
        action: AiAction = "reply",
        extraPrompt?: string,
        options?: {
            customToneId?: string;
            tone?: AiTone;
        },
    ) {
        if (isGenerating) return;
        setIsGenerating(true);
        // If we are starting a NEW task (not refining), clear previous history
        const isRefining = action === "rewrite" || action === "refine";
        const resolvedAction: AiAction = (action === "reply" || action === "forward") ? selectedAction : action;
        const rawPrompt = extraPrompt || (action === "refine" ? refineInput : prompt);
        let finalPrompt = rawPrompt;

        if (!finalPrompt) {
            if (resolvedAction === "forward") {
                finalPrompt = "Cria um email novo profissional baseado no email atualmente aberto, no contexto do caso e nos anexos selecionados. Nao trates isto como resposta ao remetente nem como exigencia de email-alvo; redige o corpo pronto a enviar para destinatarios que o utilizador escolher manualmente.";
            } else if (resolvedAction === "reply") {
                finalPrompt = replyTargetEmail
                    ? "Com base neste email, nos anexos relevantes e no contexto completo do caso, cria uma resposta final para o email-alvo selecionado. Usa o email atual como atualizacao do processo e escreve a resposta pronta a enviar."
                    : "Cria uma resposta profissional pronta a enviar com base no email atual e no contexto completo do caso.";
            }
        }

        if (resolvedAction === "reply" && replyTargetEmail) {
            const targetExcerpt = String(replyTargetEmail.bodyText || "").trim() || htmlToPlainText(String(replyTargetEmail.bodyHtml || ""));
            const targetSubject = normalizeReplySubject(replyTargetEmail.subject || ctx.subject || "");
            const replyTargetInstructions = [
                "INSTRUCOES INTERNAS PARA O ALVO DA RESPOSTA:",
                "- Esta resposta deve ser orientada ao email-alvo selecionado no dossier, e nao apenas ao email atualmente aberto no Outlook.",
                `- Email-alvo assunto: ${replyTargetEmail.subject || "(sem assunto)"}`,
                `- Email-alvo remetente: ${replyTargetEmail.fromName || replyTargetEmail.fromEmail || "--"}`,
                `- Assunto a manter na resposta: ${targetSubject}`,
                targetExcerpt ? `- Conteudo relevante do email-alvo: ${targetExcerpt.slice(0, 1600)}` : "",
                "- Usa o email atual como atualizacao/contexto complementar do caso.",
                "- Nao expliques este raciocinio ao destinatario final; usa-o apenas para fundamentar a resposta.",
            ].filter(Boolean).join("\n");
            finalPrompt = [finalPrompt, replyTargetInstructions].filter(Boolean).join("\n\n");
        }

        if (resolvedAction === "forward" && replyTargetEmail) {
            const targetExcerpt = String(replyTargetEmail.bodyText || "").trim() || htmlToPlainText(String(replyTargetEmail.bodyHtml || ""));
            const forwardContext = [
                "CONTEXTO ADICIONAL DO DOSSIER PARA NOVO EMAIL:",
                "- O email-alvo selecionado serve apenas como contexto historico; nao e destinatario automatico.",
                `- Assunto do email-alvo: ${replyTargetEmail.subject || "(sem assunto)"}`,
                `- Remetente do email-alvo: ${replyTargetEmail.fromName || replyTargetEmail.fromEmail || "--"}`,
                targetExcerpt ? `- Conteudo relevante do email-alvo: ${targetExcerpt.slice(0, 1600)}` : "",
                "- Se faltar um endereco de destinatario, nao inventes. Gera o corpo do email na mesma.",
            ].filter(Boolean).join("\n");
            finalPrompt = [finalPrompt, forwardContext].filter(Boolean).join("\n\n");
        }

        if (resolvedAction === "forward" || resolvedAction === "reply") {
            const selectedRecipients = [
                draftTo.length ? `Para selecionado: ${draftTo.join("; ")}` : "",
                draftCc.length ? `Cc selecionado: ${draftCc.join("; ")}` : "",
                draftBcc.length ? `Bcc selecionado: ${draftBcc.join("; ")}` : "",
            ].filter(Boolean).join("\n");
            const recipientInstruction = [
                "DESTINATARIOS DO RASCUNHO:",
                selectedRecipients || "- Ainda nao ha destinatarios finais selecionados. Nao inventes enderecos; escreve o corpo do email de forma reutilizavel.",
                resolvedAction === "forward" ? "- FORWARD aqui significa criar uma nova mensagem comercial baseada no email aberto." : "",
            ].filter(Boolean).join("\n");
            finalPrompt = [finalPrompt, recipientInstruction].filter(Boolean).join("\n\n");
        }

        if (!isRefining && aiState.history.length > 0) {
            setAiState({ history: [] });
        }

        setOutput("");
        setDebugLog("");
        setGenerationError("");

        try {
            const bundle = await ensureContextBundle(true);
            const analyzeFiles = await resolveSelectedAnalyzeFiles();
            const effectiveBodyTextForGeneration = effectiveBodyText || htmlToPlainText(effectiveBodyHtml || "");
            const freshSettings = await getSettings();
            const generationSelectedLocale = ((aiState.locale || freshSettings.replyLanguage || "auto") as AiLocale);

            // IMPORTANT:
            // In reply/forward mode, "auto" must be sent to the server as "auto".
            // The server is responsible for detecting the predominant language of the email.
            // Do not collapse "auto" into appLanguage/readingLanguage here, otherwise Spanish/English emails
            // can incorrectly receive Portuguese replies.
            const generationEffectiveLocale: AiLocale =
                generationSelectedLocale === "auto"
                    ? "auto"
                    : generationSelectedLocale;

            const signatureLocale: AiLocale =
                generationEffectiveLocale === "auto"
                    ? ((freshSettings.replyLanguage && freshSettings.replyLanguage !== "auto"
                        ? freshSettings.replyLanguage
                        : (freshSettings.appLanguage || "pt-PT")) as AiLocale)
                    : generationEffectiveLocale;

            const generationTone = options?.tone || aiState.tone || freshSettings.tone || "neutro";
            const effectiveCustomToneId = typeof options?.customToneId === "string"
                ? options.customToneId
                : selectedCustomToneId;
            const freshCustomTone = (freshSettings.aiCustomTones || []).find((entry: any) => entry.id === effectiveCustomToneId) || null;
            const knowledge = [...(freshSettings.aiKnowledge || [])];

            if (freshCustomTone?.instructions) {
                knowledge.push(`[TOM PERSONALIZADO ATIVO] ${String(freshCustomTone.instructions).trim()}`);
            }
            const replyDirection: AiReplyDirection | null = null;
            const signature = resolvedAction === "reply" ? buildAiSignaturePayload(freshSettings, signatureLocale) : null;
            const greetingName = resolvedAction === "reply"
                ? String(replyTargetEmail?.fromName || ctx.fromName || "").trim()
                : "";
            const greetingEmail = resolvedAction === "reply"
                ? String(replyTargetEmail?.fromEmail || ctx.fromEmail || "").trim()
                : "";

            const res = await aiGenerate({
                action: resolvedAction,
                mode: "quality",
                tone: generationTone,
                locale: generationEffectiveLocale,
                length: freshSettings.length || "m",
                inputText: finalPrompt,
                files: analyzeFiles,
                briefing: briefing, // Pass the thread summary for isolation
                contextBundle: bundle?.promptContext || "",
                email: {
                    subject: ctx.subject || "",
                    from: ctx.fromEmail || "",
                    fromName: String(ctx.fromName || "").trim(),
                    fromEmail: String(ctx.fromEmail || "").trim(),
                    greetingName,
                    greetingEmail,
                    to: (ctx.toRecipients || []).map((r: any) => r.email),
                    cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                    bodyText: effectiveBodyTextForGeneration,
                    bodyScope: freshSettings.bodyScope || "main"
                },
                persona: {
                    userRole: freshSettings.userRole,
                    styleContext: freshSettings.styleContext,
                    styleExamples: freshSettings.styleExamples,
                },
                history: isRefining ? aiState.history : [],
                knowledge,
                aiKnowledge: freshSettings.aiKnowledge || [],
                signature,
                replyDirection,
                contactAliases: freshSettings.contactAliases || [],
                // For refine: send the current editor content as explicit draft
                draftText: action === "refine" ? (output || aiState.output || "") : undefined,
            }); //inputText is already extraPrompt || prompt

            if (res.ok) {
                const formattedText = resolvedAction === "reply"
                    ? appendOfficialSignature(formatEmailHtml(res.text), signature)
                    : formatEmailHtml(res.text);
                setAiState({
                    action: resolvedAction,
                    output: formattedText,
                    suggestedTo: res.suggestedRecipients?.to || [],
                    suggestedCc: res.suggestedRecipients?.cc || [],
                    suggestedSubject: res.suggestedSubject || ""
                });
                const fullText = formattedText;
                if (looksLikeHtml(fullText)) {
                    // Rendering partial HTML fragments during the typewriter effect can corrupt
                    // the preview pane in Outlook WebView and blank the add-in surface.
                    setOutput(fullText);
                } else {
                    let current = "";
                    const words = fullText.split(/(\s+)/);
                    for (let i = 0; i < words.length; i++) {
                        current += words[i];
                        setOutput(current);
                        await new Promise((resolve) => setTimeout(resolve, 20));
                    }
                }

                const newHistory = [
                    ...(isRefining ? aiState.history : []),
                    { role: "user" as const, content: finalPrompt },
                    { role: "assistant" as const, content: fullText }
                ].slice(-4);
                setAiState({ output: fullText, history: newHistory });
                setPrompt("");
                setShowDraftPreview(true);
                setDraftDetailsExpanded(true);

                // Persist to Local History
                const entry: HistoryEntry = {
                    id: Math.random().toString(36).substring(7),
                    emailKey,
                    ts: Date.now(),
                    output: fullText,
                    prompt: finalPrompt,
                    action: resolvedAction,
                    tone: generationTone,
                    locale: generationSelectedLocale,
                    draftTo: normalizeEmailListInput(draftTo),
                    draftCc: normalizeEmailListInput(draftCc),
                    draftBcc: normalizeEmailListInput(draftBcc),
                    draftSubject: buildTicketEmailSubject(draftSubject, draftTicketCode, freshSettings?.groupTicketUi?.includeTicketCodeInSubject !== false),
                    customToneId: effectiveCustomToneId || undefined,
                    replyTarget: replyTargetEmail,
                    replyDirection,
                };
                const fullHist = [entry, ...loadHistory()];
                const savedHistory = saveHistory(fullHist);
                setHistory(savedHistory.filter(h => h.emailKey === emailKey));
            } else {
                const message = String(res.error || "Erro ao gerar resposta.");
                setGenerationError(message);
                setMsg(message);
            }
        } catch (e: any) {
            const message = String(e?.message || "Erro ao gerar resposta.");
            setGenerationError(message);
            setMsg(message);
        } finally {
            setIsGenerating(false);
        }
    }

    async function handleInsert() {
        console.log("[AiCockpit] handleInsert called");
        setDebugLog("Botão clicado. A verificar modo...");
        try {
            const isCompose = await isComposeMode();
            const includeTicketCodeInSubject = settings?.groupTicketUi?.includeTicketCodeInSubject !== false;
            const finalDraftSubject = buildTicketEmailSubject(draftSubject, draftTicketCode, includeTicketCodeInSubject);
            const forwardFiles = await resolveSelectedForwardFiles();
            console.log("[AiCockpit] isComposeMode:", isCompose);
            setDebugLog(`Modo Edição: ${isCompose}`);

            if (isCompose) {
                setDebugLog("A atualizar rascunho...");

                // Sync metadata first
                await setRecipients("to", draftTo);
                await setRecipients("cc", draftCc);
                await setRecipients("bcc", draftBcc).catch((error) => {
                    console.warn("[AiCockpit] Could not set Bcc recipients in compose:", error);
                });
                await setSubjectInComposeDraft(finalDraftSubject, { attempts: 2, delayMs: 150 });

                // Insert body
                await insertTextToBody(output);

                const composeDraftCategories = await loadDraftLinkCategories().catch(() => null);
                if (composeDraftCategories) {
                    await syncLinkCategoriesToComposeDraft(composeDraftCategories, { attempts: 1, delayMs: 0 }).catch(() => {
                        // best-effort
                    });
                }

                // Silently log for learning
                logLearningInteraction({
                    conversationId: ctx.conversationId,
                    fromEmail: ctx.fromEmail,
                    toEmails: (ctx.toRecipients || []).map((r: any) => r.email),
                    originalSubject: ctx.subject,
                    originalBody: bodyText,
                    userResponse: output
                }).catch(e => console.warn("[AiCockpit] Learning log failed:", e));

                setDebugLog("Atualizado com sucesso!");
                setMsg("Draft atualizado com sucesso!");
                setTimeout(() => setMsg(""), 3000);
                return;
            }

            // If not in compose mode, try to open a Draft based on action
            setDebugLog("A abrir rascunho (não é modo edição)...");
            const effectiveAction = selectedAction;
            const isCurrentReplyTarget = isSameStoredEmailTarget(ctx, replyTargetEmail);
            const draftLinkCategoriesPromise = loadDraftLinkCategories();
            const queueDraftCategorySync = () => {
                void draftLinkCategoriesPromise
                    .then((draftCategories) => {
                        if (!draftCategories) return;
                        return syncLinkCategoriesToComposeDraft(draftCategories);
                    })
                    .catch((error) => {
                        console.warn("[AiCockpit] Could not sync managed categories to draft:", error);
                    });
            };

            if (effectiveAction === "forward") {
                if (forwardFiles.length > 0 && (!replyTargetEmail || isCurrentReplyTarget)) {
                    await displayForwardForm(output, true);
                    try {
                        await new Promise((resolve) => setTimeout(resolve, 800));
                        if (finalDraftSubject) await setSubjectInComposeDraft(finalDraftSubject);
                        await setRecipients("to", draftTo);
                        await setRecipients("cc", draftCc);
                        await setRecipients("bcc", draftBcc).catch((error) => {
                            console.warn("[AiCockpit] Could not set Bcc recipients in forward:", error);
                        });
                    } catch (subjectError) {
                        console.warn("[AiCockpit] Could not update forward draft metadata:", subjectError);
                    }
                    queueDraftCategorySync();
                    const usedAllOriginals = forwardFiles.length === availableAttachmentCount;
                    setMsg(
                        usedAllOriginals
                            ? `Reencaminhamento aberto com ${forwardFiles.length} anexo(s) original(is).`
                            : "Reencaminhamento aberto com os anexos originais do email. Remove manualmente os que não quiseres enviar."
                    );
                    setTimeout(() => setMsg(""), 4000);
                    return;
                }

                const forwardSubject = String(draftSubject || replyTargetEmail?.subject || ctx.subject || "").trim() || "Fwd";
                await displayNewMessageForm({
                    toRecipients: draftTo,
                    ccRecipients: draftCc,
                    bccRecipients: draftBcc,
                    subject: buildTicketEmailSubject(forwardSubject, draftTicketCode, includeTicketCodeInSubject),
                    body: output,
                    isHtml: true,
                });
                queueDraftCategorySync();
                if (forwardFiles.length) {
                    try {
                        await new Promise((resolve) => setTimeout(resolve, 900));
                        for (const attachment of forwardFiles) {
                            await addBase64AttachmentToCompose(attachment.name, attachment.content);
                        }
                        setMsg(`${forwardFiles.length} anexo(s) adicionados ao rascunho.`);
                    } catch (attachError) {
                        console.warn("[AiCockpit] Could not attach selected forward files automatically:", attachError);
                        setMsg("Draft de reencaminhamento aberto. O Outlook pode exigir validação manual dos anexos nesta ação.");
                    }
                } else if (replyTargetEmail && !isCurrentReplyTarget) {
                    setMsg("Rascunho criado para o email guardado selecionado.");
                }
            } else {
                if (replyTargetEmail && !isCurrentReplyTarget) {
                    await displayNewMessageForm({
                        toRecipients: draftTo.length ? draftTo : (replyTargetEmail.fromEmail ? [replyTargetEmail.fromEmail] : []),
                        ccRecipients: draftCc,
                        bccRecipients: draftBcc,
                        subject: buildTicketEmailSubject(
                            normalizeReplySubject(draftSubject || replyTargetEmail.subject || ctx.subject || ""),
                            draftTicketCode,
                            includeTicketCodeInSubject,
                        ),
                        body: output,
                        isHtml: true,
                    });
                    queueDraftCategorySync();
                    if (forwardFiles.length) {
                        try {
                            await new Promise((resolve) => setTimeout(resolve, 900));
                            for (const attachment of forwardFiles) {
                                await addBase64AttachmentToCompose(attachment.name, attachment.content);
                            }
                            setMsg(`${forwardFiles.length} anexo(s) adicionados ao rascunho de resposta.`);
                        } catch (attachError) {
                            console.warn("[AiCockpit] Could not attach selected reply files automatically:", attachError);
                            setMsg("Rascunho criado para o email selecionado. O Outlook pode exigir validacao manual dos anexos.");
                        }
                    } else {
                        setMsg("Rascunho criado para o email guardado selecionado.");
                    }
                } else {
                    // Default to Reply (including for refine, rewrite, etc.)
                    await displayReplyForm(output);
                    if (finalDraftSubject) {
                        try {
                            await new Promise((resolve) => setTimeout(resolve, 800));
                            await setSubjectInComposeDraft(finalDraftSubject);
                        } catch (subjectError) {
                            console.warn("[AiCockpit] Could not update reply draft subject with ticket code:", subjectError);
                        }
                    }
                    queueDraftCategorySync();
                }
            }
            setDebugLog("Janela aberta.");
        } catch (e: any) {
            console.error("[AiCockpit] handleInsert error:", e);
            setDebugLog(`Erro: ${e.message}`);
            setMsg(`Erro ao inserir: ${e.message}`);
        }
    }

    const handleExport = () => {
        if (!output) return;
        const blob = new Blob([output], { type: "text/plain" });
        const url = URL.createObjectURL(blob);
        const a = document.createElement("a");
        a.href = url;
        a.download = `ia_resposta_${new Date().toISOString().slice(0, 10)}.txt`;
        document.body.appendChild(a);
        a.click();
        document.body.removeChild(a);
        URL.revokeObjectURL(url);
    };

    const activeMenu: QuickPanelId = null;
    const setActiveMenu = (_next: QuickPanelId) => { /* legacy dropdowns disabled in favor of inline quick panels */ };
    const menuRef = useRef<HTMLDivElement>(null);

    const handleKeyDown = (e: React.KeyboardEvent, action: AiAction = "reply") => {
        // Alt + Enter: New Line (allow default)
        if (e.key === "Enter" && e.altKey) {
            return;
        }
        // Enter: Send
        if (e.key === "Enter" && !e.shiftKey) {
            e.preventDefault();
            handleGenerate(action);
        }
        if (e.key === "Escape") {
            setPrompt("");
        }
    };

    const handleResetConversation = () => {
        setAiState({
            history: [],
            output: "",
            smartReplies: [],
            suggestedTo: [],
            suggestedCc: [],
            suggestedSubject: "",
            action: "reply",
        });
        setPrompt("");
        setOutput("");
        setBriefing(null);
        setExtractedTasks([]);
        setShowTaskReview(false);
        setDraftTo([]);
        setDraftCc([]);
        setDraftBcc([]);
        setDraftSubject("");
        setShowDraftPreview(false);
        setDraftDetailsExpanded(false);
        setSuggestedContacts([]);
        setSelectedCustomToneId("");
        setActivePanel(null);
        setHistoryExpanded(false);
        setReplyTargetEmail(null);
        setReplyAddresseeName("");
        setReplyAddresseeContext("");
        clearFiles();
    };

    const toneRefiners: Array<{ label: string; tone: AiTone; icon: React.ReactNode }> = [
        { label: "Prof.", tone: "formal", icon: <Icons.Building size={12} /> },
        { label: "Amig.", tone: "simpático", icon: <Icons.Sparkles size={12} /> },
        { label: "Curto", tone: "curto", icon: <Icons.Receipt size={12} /> },
    ];

    const MiniFlag: React.FC<{ locale: string }> = ({ locale }) => {
        if (locale === "auto") return <span>🤖</span>;
        const flags: Record<string, React.ReactNode> = {
            "pt-PT": (
                <svg viewBox="0 0 512 512" width="100%" height="100%">
                    <rect width="204.8" height="512" fill="#006600" />
                    <rect x="204.8" width="307.2" height="512" fill="#ff0000" />
                    <circle cx="204.8" cy="256" r="102.4" fill="#ffff00" opacity="0.8" />
                </svg>
            ),
            "en-GB": (
                <svg viewBox="0 0 512 512" width="100%" height="100%">
                    <rect width="512" height="512" fill="#00247d" />
                    <path d="M0 0l512 512M512 0L0 512" stroke="#fff" strokeWidth="60" />
                    <path d="M0 0l512 512M512 0L0 512" stroke="#cf142b" strokeWidth="40" />
                    <path d="M256 0v512M0 256h512" stroke="#fff" strokeWidth="100" />
                    <path d="M256 0v512M0 256h512" stroke="#cf142b" strokeWidth="60" />
                </svg>
            ),
            "es-ES": (
                <svg viewBox="0 0 512 512" width="100%" height="100%">
                    <rect width="512" height="128" fill="#c60b1e" />
                    <rect y="128" width="512" height="256" fill="#ffc400" />
                    <rect y="384" width="512" height="128" fill="#c60b1e" />
                </svg>
            ),
            "it-IT": (
                <svg viewBox="0 0 512 512" width="100%" height="100%">
                    <rect width="170.7" height="512" fill="#009246" />
                    <rect x="170.7" width="170.7" height="512" fill="#fff" />
                    <rect x="341.4" width="170.7" height="512" fill="#ce2b37" />
                </svg>
            ),
            "de-DE": (
                <svg viewBox="0 0 512 512" width="100%" height="100%">
                    <rect width="512" height="170.7" fill="#000" />
                    <rect y="170.7" width="512" height="170.7" fill="#d00" />
                    <rect y="341.4" width="512" height="170.7" fill="#ffce00" />
                </svg>
            )
        };
        return flags[locale] || <span>🏳️</span>;
    };
    const localeOptions: Array<{ label: string; value: AiLocale }> = [
        { label: "Auto", value: "auto" },
        { label: "Português", value: "pt-PT" },
        { label: "English", value: "en-GB" },
        { label: "Español", value: "es-ES" },
        { label: "Italiano", value: "it-IT" },
        { label: "Deutsch", value: "de-DE" },
    ];

    const S: Record<string, React.CSSProperties> = {
        primaryBtnPill: {
            boxSizing: "border-box",
            height: px(22), minHeight: px(22), maxHeight: px(22),
            borderRadius: px(14),
            border: "1px solid rgba(0, 80, 180, 0.4)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: "4px",
            padding: `0 ${px(8)}`,
            fontSize: fpx(9),
            fontWeight: 400,
            lineHeight: 1,
            textTransform: "uppercase",
            cursor: "pointer",
            outline: "none",
            background: "linear-gradient(180deg, rgba(80, 160, 255, 0.95) 0%, rgba(0, 100, 210, 0.85) 100%)",
            color: "#FFFFFF",
            boxShadow: "0 4px 10px rgba(0,100,210,0.35), inset 0 1px 0 rgba(255,255,255,0.55), inset 0 -1px 0 rgba(0,0,0,0.15)",
            transition: "all 0.18s ease",
        },
        secondaryBtnPill: {
            boxSizing: "border-box",
            width: px(24), minWidth: px(24), maxWidth: px(24),
            height: px(22), minHeight: px(22), maxHeight: px(22),
            borderRadius: px(14),
            border: "1px solid rgba(200, 210, 230, 0.6)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: "5px",
            padding: "0",
            fontSize: fpx(10),
            fontWeight: 400,
            lineHeight: 1,
            textTransform: "uppercase",
            cursor: "pointer",
            outline: "none",
            background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
            color: "#172B4D",
            boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
            transition: "all 0.18s ease",
        },
        secondaryBtnLink: {
            boxSizing: "border-box",
            height: px(22), minHeight: px(22), maxHeight: px(22),
            borderRadius: "16px",
            border: "1px solid rgba(200, 210, 230, 0.6)",
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            gap: "5px",
            padding: `0 ${px(6)}`,
            fontSize: fpx(9),
            fontWeight: 400,
            lineHeight: 1,
            textTransform: "uppercase",
            cursor: "pointer",
            outline: "none",
            background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
            color: "#172B4D",
            boxShadow: "0 4px 10px rgba(0,0,0,0.1), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06)",
            transition: "all 0.18s ease",
        },
        cascadeMenu: {
            position: "absolute",
            top: "calc(100% + 6px)",
            bottom: "auto",
            left: 0,
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            zIndex: 9999,
            background: "transparent",
            width: "fit-content"
        },
        cascadeItem: {
            boxSizing: "border-box",
            height: "auto",
            minHeight: px(22),
            minWidth: px(72),
            whiteSpace: "normal",
            borderRadius: "16px",
            border: "1px solid rgba(200, 210, 230, 0.6)",
            backdropFilter: "blur(12px)",
            WebkitBackdropFilter: "blur(12px)",
            display: "flex",
            alignItems: "center",
            justifyContent: "flex-start",
            gap: "6px",
            padding: `4px ${px(8)}`,
            fontSize: fpx(9),
            fontWeight: 400,
            lineHeight: 1.2,
            textTransform: "uppercase",
            cursor: "pointer",
            background: "linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(240,245,255,0.95) 100%)",
            color: "#172B4D",
            boxShadow: "0 4px 12px rgba(0,0,0,0.12), inset 0 1px 0 rgba(255,255,255,1)",
            transition: "all 0.18s ease",
            textAlign: "left"
        },
        container: {
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            paddingTop: "2px",
        },
        inputCard: {
            background: "var(--iccc-card-bg)",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "10px",
            padding: "8px 10px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.02)",
            display: "flex",
            flexDirection: "column",
            gap: "4px",
        },
        textarea: {
            width: "100%",
            minHeight: "70px",
            background: "transparent",
            border: "none",
            color: "var(--iccc-text)",
            fontFamily: "var(--iccc-font)",
            fontSize: "12px",
            resize: "none",
            outline: "none",
        },
        inputFooter: {
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
        },
        refinerRow: {
            display: "flex",
            gap: isNarrow ? "2px" : "4px",
            paddingBottom: "4px",
            alignItems: "center",
            overflow: "visible",
            maxWidth: "100%",
            overflowX: "visible",
            flexWrap: "nowrap",
            position: "relative",
            zIndex: 5
        },
        outputCard: {
            background: "var(--iccc-card-bg)",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "12px",
            padding: "12px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.03)",
            display: "flex",
            flexDirection: "column",
            gap: "8px",
            animation: "fadeIn 0.3s ease",
        },
        outputHeader: {
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            padding: "6px 0",
            fontSize: "10px",
            fontWeight: 400,
            textTransform: "uppercase",
            letterSpacing: "0.5px",
            color: "var(--iccc-text-muted)",
        },
        outputText: {
            fontSize: "13px",
            lineHeight: "1.25",
            color: "var(--iccc-text)",
            whiteSpace: "pre-wrap",
        },
        typingDots: {
            display: "flex",
            gap: "3px",
        },
        intentContainer: {
            display: "flex",
            flexWrap: "wrap" as const,
            gap: "6px",
            padding: "0 4px",
            marginBottom: "2px",
        },
        intentChip: {
            background: "var(--iccc-card-bg)",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "12px",
            padding: "3px 8px",
            fontSize: "10px",
            color: "var(--iccc-text)",
            cursor: "pointer",
            transition: "all 0.2s ease",
            whiteSpace: "nowrap" as const,
            boxShadow: "0 1px 2px rgba(0,0,0,0.03)",
        },
        skeletonText: {
            fontSize: "11px",
            color: "var(--iccc-text-muted)",
            fontStyle: "italic",
            padding: "4px 10px",
            display: "flex",
            alignItems: "center",
            gap: "6px",
        },
        chatInputWrapper: {
            display: "flex",
            alignItems: "center",
            gap: "6px",
            background: "var(--iccc-bg)",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "10px",
            padding: "2px 2px 2px 10px",
        },
        chatInput: {
            flex: 1,
            background: "transparent",
            border: "none",
            color: "var(--iccc-text)",
            fontSize: "12px",
            outline: "none",
            padding: "6px 0",
        },
        chatSendBtn: {
            background: "var(--iccc-btn-bg)",
            color: "white",
            border: "none",
            borderRadius: "8px",
            width: px(28),
            height: px(22),
            display: "flex",
            alignItems: "center",
            justifyContent: "center",
            cursor: "pointer",
        },
        actionBtn: {
            background: "none",
            border: "none",
            color: "var(--iccc-text-muted)",
            fontSize: "11px",
            fontWeight: 600,
            cursor: "pointer",
        },
        actionBtnPrimary: {
            background: "none",
            border: "none",
            color: "#3b82f6",
            fontSize: "11px",
            fontWeight: 600,
            cursor: "pointer",
        },
        briefingCard: {
            background: "rgba(59, 130, 246, 0.05)",
            border: "1px solid rgba(59, 130, 246, 0.12)",
            borderRadius: "10px",
            padding: "8px 10px",
            marginBottom: "6px",
            display: "flex",
            flexDirection: "column",
            gap: "4px",
            position: "relative",
            transition: "all 0.3s ease",
        },
        briefingHeader: {
            display: "flex",
            justifyContent: "space-between",
            alignItems: "center",
            fontSize: "10px",
            fontWeight: 400,
            color: "#1e40af",
            textTransform: "uppercase",
            letterSpacing: "0.5px",
        },
        briefingContent: {
            fontSize: "11px",
            lineHeight: "1.4",
            color: "#334155",
            overflow: "hidden",
            display: "-webkit-box",
            WebkitBoxOrient: "vertical",
            transition: "all 0.3s ease",
        },
        draftCard: {
            background: "#f8fafc",
            border: "1px solid #e2e8f0",
            borderRadius: "8px",
            margin: "4px 0",
            overflow: "hidden",
        },
        draftHeader: {
            background: "#f1f5f9",
            padding: "6px 10px",
            fontSize: "10px",
            fontWeight: 400,
            color: "#64748b",
            display: "flex",
            alignItems: "center",
            gap: "6px",
            cursor: "pointer",
            textTransform: "uppercase",
        },
        draftBody: {
            padding: "8px 10px",
            display: "flex",
            flexDirection: "column",
            gap: "6px",
        },
        draftRow: {
            display: "flex",
            alignItems: "center",
            gap: "8px",
        },
        draftLabel: {
            fontSize: "10px",
            fontWeight: 700,
            color: "#94a3b8",
            width: "50px",
            flexShrink: 0,
        },
        draftInput: {
            flex: 1,
            background: "#fff",
            border: "1px solid #e2e8f0",
            borderRadius: "4px",
            fontSize: "11px",
            padding: "2px 6px",
            color: "#1e293b",
            outline: "none",
        },
        suggestedChip: {
            background: "#dbeafe",
            color: "#1e40af",
            border: "1px solid #bfdbfe",
            borderRadius: "12px",
            padding: "2px 8px",
            fontSize: "10px",
            fontWeight: 600,
            cursor: "pointer",
            transition: "all 0.2s ease",
        },
        actionToggleRow: {
            display: "flex",
            gap: "6px",
            alignItems: "center",
            marginBottom: "4px",
        },
        actionToggleBtn: {
            boxSizing: "border-box",
            height: px(22),
            borderRadius: "999px",
            border: "1px solid rgba(200, 210, 230, 0.6)",
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            padding: `0 ${px(10)}`,
            fontSize: fpx(9),
            fontWeight: 600,
            letterSpacing: "0.04em",
            textTransform: "uppercase",
            background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
            color: "#475569",
            cursor: "pointer",
        },
        actionToggleBtnOn: {
            boxSizing: "border-box",
            height: px(22),
            borderRadius: "999px",
            border: "1px solid rgba(0, 80, 180, 0.4)",
            display: "inline-flex",
            alignItems: "center",
            justifyContent: "center",
            padding: `0 ${px(10)}`,
            fontSize: fpx(9),
            fontWeight: 700,
            letterSpacing: "0.04em",
            textTransform: "uppercase",
            background: "linear-gradient(180deg, rgba(80, 160, 255, 0.95) 0%, rgba(0, 100, 210, 0.85) 100%)",
            color: "#FFFFFF",
            cursor: "pointer",
        },
        replyTargetRow: {
            display: "flex",
            alignItems: "stretch",
            gap: "6px",
            marginBottom: "4px",
        },
        replyTargetBtn: {
            ...{
                boxSizing: "border-box",
                height: px(22),
                borderRadius: "999px",
                border: "1px solid rgba(200, 210, 230, 0.6)",
                display: "inline-flex",
                alignItems: "center",
                justifyContent: "center",
                padding: `0 ${px(10)}`,
                fontSize: fpx(9),
                fontWeight: 700,
                letterSpacing: "0.04em",
                textTransform: "uppercase",
                background: "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
                color: "#475569",
                cursor: "pointer",
            },
            gap: "6px",
            flexShrink: 0,
        },
        replyTargetSummary: {
            flex: 1,
            minWidth: 0,
            borderRadius: "12px",
            border: "1px solid rgba(37, 99, 235, 0.14)",
            background: "rgba(239, 246, 255, 0.9)",
            padding: "6px 8px",
            display: "grid",
            gap: "2px",
        },
        replyTargetTitle: {
            fontSize: "10px",
            fontWeight: 800,
            color: "#1d4ed8",
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
        },
        replyTargetMeta: {
            fontSize: "10px",
            color: "#475569",
            overflow: "hidden",
            textOverflow: "ellipsis",
            whiteSpace: "nowrap",
        },
        replyTargetClear: {
            border: "none",
            background: "transparent",
            color: "#64748b",
            fontSize: "11px",
            cursor: "pointer",
            padding: "0 2px",
            alignSelf: "flex-start",
        },
        quickPanel: {
            background: "var(--iccc-card-bg)",
            border: "1px solid var(--iccc-card-border)",
            borderRadius: "12px",
            padding: "8px",
            display: "grid",
            gap: "8px",
            boxShadow: "0 1px 4px rgba(0,0,0,0.03)",
        },
        quickPanelHeader: {
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: "8px",
            fontSize: "10px",
            fontWeight: 700,
            color: "#475569",
            textTransform: "uppercase",
            letterSpacing: "0.04em",
        },
        quickPanelBody: {
            display: "grid",
            gap: "6px",
            maxHeight: "240px",
            overflowY: "auto",
            paddingRight: "2px",
        },
        quickPanelSearch: {
            width: "100%",
            background: "#fff",
            border: "1px solid #dbe3f3",
            borderRadius: "10px",
            fontSize: "11px",
            color: "#172B4D",
            outline: "none",
            padding: "6px 10px",
        },
        quickPanelItem: {
            border: "1px solid #dbe3f3",
            background: "#fff",
            borderRadius: "10px",
            padding: "8px 10px",
            display: "flex",
            alignItems: "center",
            justifyContent: "space-between",
            gap: "8px",
            fontSize: "11px",
            color: "#172B4D",
        },
        quickPanelItemBtn: {
            border: "1px solid #dbe3f3",
            background: "#fff",
            borderRadius: "10px",
            padding: "8px 10px",
            display: "flex",
            alignItems: "center",
            gap: "8px",
            fontSize: "11px",
            color: "#172B4D",
            cursor: "pointer",
            textAlign: "left",
        },
        purposeChip: {
            border: "1px solid #dbe3f3",
            background: "#fff",
            borderRadius: "999px",
            padding: "3px 8px",
            fontSize: "9px",
            color: "#475569",
            cursor: "pointer",
        },
        purposeChipOn: {
            border: "1px solid rgba(37, 99, 235, 0.2)",
            background: "rgba(37, 99, 235, 0.08)",
            borderRadius: "999px",
            padding: "3px 8px",
            fontSize: "9px",
            color: "#1d4ed8",
            cursor: "pointer",
        },
    };

    const filteredPresets = (settings?.responsePresets || [])
        .filter((p: any) => {
            const name = String(p?.name || "").toLowerCase();
            const promptText = String(p?.prompt || "").toLowerCase();
            const search = presetSearch.toLowerCase();
            return !search || name.includes(search) || promptText.includes(search);
        })
        .slice(0, 12);

    const filteredIntents = aiState.smartReplies
        .filter((intent: string) => !intentSearch || intent.toLowerCase().includes(intentSearch.toLowerCase()));

    const filteredContacts = [
        ...suggestedContacts
            .filter((email) => !contactSearch || email.toLowerCase().includes(contactSearch.toLowerCase()))
            .map((email) => ({ kind: "email" as const, id: email, label: email, value: email })),
        ...((settings?.contactAliases || [])
            .filter((entry: any) => !contactSearch || entry.name.toLowerCase().includes(contactSearch.toLowerCase()) || entry.email.toLowerCase().includes(contactSearch.toLowerCase()))
            .map((entry: any) => ({ kind: "alias" as const, id: entry.id, label: entry.name, value: entry.email }))),
    ];

    const draftRecipientSuggestions = useMemo(() => normalizeEmailListInput([
        ctx.fromEmail,
        ...(ctx.toRecipients || []).map((recipient: any) => recipient?.email),
        ...(ctx.ccRecipients || []).map((recipient: any) => recipient?.email),
        replyTargetEmail?.fromEmail,
        ...suggestedContacts,
        ...((settings?.contactAliases || []).map((entry: any) => entry?.email)),
    ]).slice(0, 12), [ctx.ccRecipients, ctx.fromEmail, ctx.toRecipients, replyTargetEmail?.fromEmail, settings?.contactAliases, suggestedContacts]);

    function addDraftRecipient(kind: "to" | "cc" | "bcc", email: string) {
        if (kind === "to") setDraftTo((prev) => addUniqueEmail(prev, email));
        if (kind === "cc") setDraftCc((prev) => addUniqueEmail(prev, email));
        if (kind === "bcc") setDraftBcc((prev) => addUniqueEmail(prev, email));
    }

    const filteredAttachments = (persistedEmailAttachments || [])
        .filter((entry) => {
            const name = String(entry?.name || "").trim().toLowerCase();
            return !fileSearch || name.includes(fileSearch.toLowerCase());
        });

    function renderQuickPanel() {
        if (!activePanel) return null;

        const closeButton = (
            <button
                type="button"
                style={{ ...S.actionBtn, fontSize: "10px" }}
                onClick={() => setActivePanel(null)}
            >
                Fechar
            </button>
        );

        if (activePanel === "lang") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>Idioma</span>
                        {closeButton}
                    </div>
                    <div style={S.quickPanelBody}>
                        {localeOptions.map((opt) => (
                            <button
                                key={opt.value}
                                type="button"
                                style={S.quickPanelItemBtn}
                                onClick={() => {
                                    setAiState({ locale: opt.value });
                                    setActivePanel(null);
                                    if (output) void handleGenerate("rewrite", output);
                                }}
                            >
                                <div style={{ width: "16px", height: "16px", borderRadius: "50%", overflow: "hidden", display: "flex", alignItems: "center", justifyContent: "center", background: "rgba(0,0,0,0.05)" }}>
                                    <MiniFlag locale={opt.value} />
                                </div>
                                <div style={{ display: "grid", gap: "2px" }}>
                                    <span style={{ fontSize: "11px", fontWeight: 700 }}>{opt.label}</span>
                                    <span style={{ fontSize: "9px", color: "#64748b" }}>{opt.value === selectedLocale ? "Ativo" : "Selecionar"}</span>
                                </div>
                            </button>
                        ))}
                    </div>
                </div>
            );
        }

        if (activePanel === "presets") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>MODS</span>
                        {closeButton}
                    </div>
                    <input
                        style={S.quickPanelSearch}
                        placeholder="Procurar modelo..."
                        value={presetSearch}
                        onChange={(e) => setPresetSearch(e.target.value)}
                    />
                    <div style={S.quickPanelBody}>
                        {filteredPresets.map((preset: any) => (
                            <button
                                key={preset.id}
                                type="button"
                                style={S.quickPanelItemBtn}
                                onClick={() => {
                                    setActivePanel(null);
                                    setPresetSearch("");
                                    const presetPrompt = String(preset?.prompt || "").trim();
                                    const taskInstruction = presetPrompt
                                        ? `MOD selecionado: ${String(preset?.name || "MOD").trim()}\n\nInstrucao obrigatoria do MOD:\n${presetPrompt}`
                                        : "";
                                    void handleGenerate(selectedAction, taskInstruction);
                                }}
                            >
                                <Icons.ArrowRight size={12} />
                                <div style={{ display: "grid", gap: "2px", flex: 1 }}>
                                    <span style={{ fontSize: "11px", fontWeight: 700 }}>{preset.name}</span>
                                    <span style={{ fontSize: "9px", color: "#64748b", whiteSpace: "normal" }}>{preset.prompt}</span>
                                </div>
                            </button>
                        ))}
                        {!filteredPresets.length ? <div style={{ ...S.quickPanelItem, color: "#64748b" }}>Sem modelos configurados.</div> : null}
                    </div>
                </div>
            );
        }

        if (activePanel === "intents") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>Dicas</span>
                        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                            <button type="button" style={{ ...S.actionBtn, fontSize: "10px" }} onClick={() => { void handleFetchIntents(true); }}>
                                Atualizar
                            </button>
                            {closeButton}
                        </div>
                    </div>
                    <input
                        style={S.quickPanelSearch}
                        placeholder="Procurar sugestão..."
                        value={intentSearch}
                        onChange={(e) => setIntentSearch(e.target.value)}
                    />
                    <div style={S.quickPanelBody}>
                        {isFetchingIntents ? <div style={{ ...S.quickPanelItem, color: "#64748b" }}>A gerar sugestões...</div> : null}
                        {!isFetchingIntents && filteredIntents.map((intent: string, idx: number) => (
                            <button
                                key={`${intent}-${idx}`}
                                type="button"
                                style={S.quickPanelItemBtn}
                                onClick={() => {
                                    setPrompt(intent);
                                    setActivePanel(null);
                                    setIntentSearch("");
                                    void handleGenerate(selectedAction, intent);
                                }}
                            >
                                <Icons.Sparkles size={12} />
                                <span style={{ fontSize: "11px", fontWeight: 700, whiteSpace: "normal" }}>{intent}</span>
                            </button>
                        ))}
                        {!isFetchingIntents && !filteredIntents.length ? <div style={{ ...S.quickPanelItem, color: "#64748b" }}>Sem sugestões disponíveis.</div> : null}
                    </div>
                </div>
            );
        }

        if (activePanel === "contacts") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>Destinatários</span>
                        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                            <button type="button" style={{ ...S.actionBtn, fontSize: "10px" }} onClick={() => { void handleExtractContacts(true); }}>
                                Atualizar
                            </button>
                            {closeButton}
                        </div>
                    </div>
                    <input
                        style={S.quickPanelSearch}
                        placeholder="Procurar contacto..."
                        value={contactSearch}
                        onChange={(e) => setContactSearch(e.target.value)}
                    />
                    <div style={S.quickPanelBody}>
                        {filteredContacts.map((entry) => (
                            <div key={`${entry.kind}-${entry.id}`} style={S.quickPanelItem}>
                                <div style={{ display: "grid", gap: "2px", flex: 1 }}>
                                    <span style={{ fontSize: "11px", fontWeight: 700 }}>{entry.label}</span>
                                    <span style={{ fontSize: "9px", color: "#64748b" }}>{entry.value}</span>
                                </div>
                                <div style={{ display: "flex", gap: "6px" }}>
                                    <button
                                        type="button"
                                        style={draftTo.includes(entry.value) ? S.purposeChipOn : S.purposeChip}
                                        onClick={() => setDraftTo((prev) => prev.includes(entry.value) ? prev.filter((value) => value !== entry.value) : addUniqueEmail(prev, entry.value))}
                                    >
                                        To
                                    </button>
                                    <button
                                        type="button"
                                        style={draftCc.includes(entry.value) ? S.purposeChipOn : S.purposeChip}
                                        onClick={() => setDraftCc((prev) => prev.includes(entry.value) ? prev.filter((value) => value !== entry.value) : addUniqueEmail(prev, entry.value))}
                                    >
                                        Cc
                                    </button>
                                    <button
                                        type="button"
                                        style={draftBcc.includes(entry.value) ? S.purposeChipOn : S.purposeChip}
                                        onClick={() => setDraftBcc((prev) => prev.includes(entry.value) ? prev.filter((value) => value !== entry.value) : addUniqueEmail(prev, entry.value))}
                                    >
                                        Bcc
                                    </button>
                                </div>
                            </div>
                        ))}
                        {!filteredContacts.length ? <div style={{ ...S.quickPanelItem, color: "#64748b" }}>Sem contactos disponíveis.</div> : null}
                    </div>
                </div>
            );
        }

        if (activePanel === "files") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>Ficheiros</span>
                        <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                            <button
                                type="button"
                                style={{ ...S.actionBtn, fontSize: "10px" }}
                                onClick={() => {
                                    for (const attachment of persistedEmailAttachments || []) {
                                        setAttachmentUsage(String(attachment?.name || ""), { analyze: true });
                                    }
                                }}
                            >
                                Usar todos
                            </button>
                            {closeButton}
                        </div>
                    </div>
                    <input
                        style={S.quickPanelSearch}
                        placeholder="Procurar ficheiro..."
                        value={fileSearch}
                        onChange={(e) => setFileSearch(e.target.value)}
                    />
                    <div style={S.quickPanelBody}>
                        {filteredAttachments.map((attachment) => {
                            const name = String(attachment?.name || "").trim();
                            const usage = fileUsage[name] || { analyze: false, forward: false };
                            return (
                                <div key={name} style={S.quickPanelItem}>
                                    <div style={{ display: "grid", gap: "2px", flex: 1 }}>
                                        <span style={{ fontSize: "11px", fontWeight: 700 }}>{name}</span>
                                        <span style={{ fontSize: "9px", color: "#64748b" }}>
                                            {[attachment.contentType, attachment.size ? `${Math.round(Number(attachment.size) / 1024)} KB` : ""].filter(Boolean).join(" | ")}
                                        </span>
                                    </div>
                                    <div style={{ display: "flex", gap: "6px" }}>
                                        <button
                                            type="button"
                                            style={usage.analyze ? S.purposeChipOn : S.purposeChip}
                                            onClick={() => setAttachmentUsage(name, { analyze: !usage.analyze })}
                                        >
                                            Analisar
                                        </button>
                                        <button
                                            type="button"
                                            style={usage.forward ? S.purposeChipOn : S.purposeChip}
                                            onClick={() => setAttachmentUsage(name, { forward: !usage.forward })}
                                        >
                                            Reenviar
                                        </button>
                                    </div>
                                </div>
                            );
                        })}
                        {!filteredAttachments.length ? <div style={{ ...S.quickPanelItem, color: "#64748b" }}>Sem anexos disponíveis neste email.</div> : null}
                    </div>
                </div>
            );
        }

        if (activePanel === "mode") {
            return (
                <div style={S.quickPanel}>
                    <div style={S.quickPanelHeader}>
                        <span>Modo</span>
                        {closeButton}
                    </div>
                    <div style={S.quickPanelBody}>
                        {toneRefiners.map((toneOption) => (
                            <button
                                key={toneOption.tone}
                                type="button"
                                style={S.quickPanelItemBtn}
                                onClick={() => {
                                    setAiState({ tone: toneOption.tone });
                                    setSelectedCustomToneId("");
                                    setActivePanel(null);
                                    void handleGenerate(selectedAction, undefined, {
                                        customToneId: "",
                                        tone: toneOption.tone,
                                    });
                                }}
                            >
                                {toneOption.icon}
                                <div style={{ display: "grid", gap: "2px", flex: 1 }}>
                                    <span style={{ fontSize: "11px", fontWeight: 700 }}>{toneOption.label}</span>
                                    <span style={{ fontSize: "9px", color: "#64748b" }}>{baseTone === toneOption.tone && !selectedCustomToneId ? "Ativo" : "Selecionar"}</span>
                                </div>
                            </button>
                        ))}
                        {(settings?.aiCustomTones || []).map((toneEntry: any) => (
                            <button
                                key={toneEntry.id}
                                type="button"
                                style={S.quickPanelItemBtn}
                                onClick={() => {
                                    setSelectedCustomToneId(toneEntry.id);
                                    setActivePanel(null);
                                    void handleGenerate(selectedAction, undefined, {
                                        customToneId: toneEntry.id,
                                    });
                                }}
                            >
                                <Icons.Sparkles size={12} />
                                <div style={{ display: "grid", gap: "2px", flex: 1 }}>
                                    <span style={{ fontSize: "11px", fontWeight: 700 }}>{toneEntry.name}</span>
                                    <span style={{ fontSize: "9px", color: "#64748b", whiteSpace: "normal" }}>
                                        {selectedCustomToneId === toneEntry.id ? "Ativo" : String(toneEntry.instructions || "").trim() || "Tom personalizado"}
                                    </span>
                                </div>
                            </button>
                        ))}
                    </div>
                </div>
            );
        }

        return null;
    }


    return (
        <div style={S.container} ref={paneRef}>
            {/* Glossy Pill Hover Styling */}
            <style>{`
                .iccc-glossy-pill {
                    transition: all 0.18s ease !important;
                }
                .iccc-glossy-pill:hover {
                    transform: translateY(-1px);
                    filter: brightness(1.02);
                }
                .iccc-primary-pill:hover {
                    box-shadow: 0 6px 14px rgba(0,100,210,0.5), inset 0 1px 0 rgba(255,255,255,0.7), inset 0 -1px 0 rgba(0,0,0,0.15) !important;
                }
                .iccc-secondary-pill:hover {
                    background: linear-gradient(180deg, rgba(230,240,255,0.98) 0%, rgba(195,215,248,0.95) 100%) !important;
                    box-shadow: 0 6px 14px rgba(0,80,200,0.15), inset 0 1px 0 rgba(255,255,255,1), inset 0 -1px 0 rgba(0,0,0,0.06) !important;
                }
            `}</style>

            <div style={{ display: "flex", justifyContent: "flex-end", marginBottom: "6px" }}>
                <button
                    type="button"
                    style={{
                        width: "30px",
                        height: "30px",
                        borderRadius: "999px",
                        border: "1px solid rgba(148, 163, 184, 0.24)",
                        background: "rgba(255,255,255,0.86)",
                        color: "#64748b",
                        display: "inline-flex",
                        alignItems: "center",
                        justifyContent: "center",
                        cursor: "pointer",
                        boxShadow: "0 4px 12px rgba(15, 23, 42, 0.08)",
                    }}
                    onClick={() => void openAiSettings()}
                    title="Settings da IA"
                    aria-label="Abrir settings da IA"
                >
                    <Icons.Settings size={14} />
                </button>
            </div>

            {/* 30-Second Briefing Card */}
            {isFetchingBriefing && (
                <div style={{ ...S.briefingCard, background: "rgba(0,0,0,0.02)", borderStyle: "dashed" }}>
                    <div style={S.skeletonText}>
                        <Icons.RotateCcw size={10} style={{ animation: "spin 1s linear infinite" }} />
                        A gerar briefing do thread...
                    </div>
                </div>
            )}

            {!ctx.isCompose && ctx.conversationId && !isFetchingBriefing && briefing && (
                <div style={S.briefingCard}>
                    <div style={S.briefingHeader}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            <Icons.Sparkles size={12} color="#1e40af" />
                            <span>30-Second Briefing</span>
                        </div>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            <button
                                onClick={() => { void handleFetchBriefing(true); }}
                                style={{
                                    ...S.actionBtn,
                                    height: "18px",
                                    padding: "0 8px",
                                    background: "rgba(59, 130, 246, 0.08)",
                                    borderRadius: "10px",
                                    color: "#1e40af",
                                    fontSize: "9px"
                                }}
                            >
                                Atualizar
                            </button>
                            <button
                                onClick={() => setBriefingExpanded(!briefingExpanded)}
                                style={{
                                    ...S.actionBtn,
                                    height: "18px",
                                    padding: "0 8px",
                                    background: "rgba(59, 130, 246, 0.1)",
                                    borderRadius: "10px",
                                    color: "#1e40af",
                                    fontSize: "9px"
                                }}
                            >
                                {briefingExpanded ? "Recolher" : "Expandir"}
                            </button>
                        </div>
                    </div>
                    <div
                        style={{
                            ...S.briefingContent,
                            WebkitLineClamp: briefingExpanded ? "unset" : 2,
                            maxHeight: briefingExpanded ? "300px" : "34px",
                        } as any}
                    >
                        {briefing}
                    </div>
                </div>
            )}

            {aiManualOnly && !ctx.isCompose && ctx.conversationId && !isFetchingBriefing && !briefing && (
                <div style={{ ...S.briefingCard, background: "rgba(59, 130, 246, 0.04)", borderStyle: "dashed" }}>
                    <div style={S.briefingHeader}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            <Icons.Sparkles size={12} color="#1e40af" />
                            <span>30-Second Briefing</span>
                        </div>
                    </div>
                    <div style={{ ...S.briefingContent, fontSize: "10px", color: "#475569" }}>
                        O resumo deixou de ser automatico. Gera-o apenas quando precisares.
                    </div>
                    <div style={{ display: "flex", marginTop: "6px" }}>
                        <button
                            onClick={() => { void handleFetchBriefing(true); }}
                            style={{ ...S.actionBtnPrimary, fontSize: "10px" }}
                        >
                            Gerar briefing
                        </button>
                    </div>
                </div>
            )}

            {isDevRuntime && debugLog && (
                <div style={{ padding: "8px", background: "#fee2e2", color: "#b91c1c", fontSize: "11px", borderRadius: "4px", border: "1px solid #fca5a5" }}>
                    DEBUG: {debugLog}
                </div>
            )}

            {/* Intenções Sugeridas (Smart Replies) */}
            {!output && !isGenerating && (aiState.smartReplies.length > 0 || isFetchingIntents) && (
                <div style={S.intentContainer}>
                    {isFetchingIntents ? (
                        <div style={S.skeletonText}>A sugerir respostas...</div>
                    ) : (
                        <div style={{ padding: "4px 8px", fontSize: "10px", color: "var(--iccc-text-muted)" }}>
                            Usa o menu <strong>SUGESTÕES</strong> acima para respostas rápidas.
                        </div>
                    )}
                </div>
            )}

            {isExtractingTasks && (
                <div style={{ ...S.briefingCard, background: "rgba(16, 185, 129, 0.03)", borderStyle: "dashed", borderColor: "rgba(16, 185, 129, 0.2)" }}>
                    <div style={{ ...S.skeletonText, color: "#059669" }}>
                        <Icons.RotateCcw size={10} style={{ animation: "spin 1s linear infinite" }} />
                        A identificar tarefas...
                    </div>
                </div>
            )}

            {aiManualOnly && !ctx.isCompose && ctx.conversationId && !isExtractingTasks && extractedTasks.length === 0 && (
                <div style={{ ...S.briefingCard, background: "rgba(16, 185, 129, 0.03)", borderStyle: "dashed", borderColor: "rgba(16, 185, 129, 0.2)" }}>
                    <div style={{ ...S.briefingHeader, color: "#047857" }}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            <Icons.Clipboard size={12} />
                            <span>Tarefas do Email</span>
                        </div>
                    </div>
                    <div style={{ ...S.briefingContent, fontSize: "10px", color: "#065f46" }}>
                        A deteccao de tarefas passou para manual. Corre apenas quando quiseres rever acionaveis.
                    </div>
                    <div style={{ display: "flex", marginTop: "6px" }}>
                        <button
                            onClick={() => { void handleExtractTasksReview(); }}
                            style={{ ...S.actionBtnPrimary, fontSize: "10px", color: "#059669" }}
                        >
                            Detetar tarefas
                        </button>
                    </div>
                </div>
            )}

            {/* AI Task Extraction Review */}
            {extractedTasks.length > 0 && (
                <div style={{ ...S.draftCard, border: "1px solid #10b981", background: "#f0fdf4", marginBottom: "8px" }}>
                    <div style={{ ...S.draftHeader, background: "#dcfce7", color: "#065f46" }} onClick={() => setShowTaskReview(!showTaskReview)}>
                        <Icons.Clipboard size={12} />
                        <span style={{ flex: 1 }}>Tarefas Detetadas ({extractedTasks.length})</span>
                        <Icons.ArrowDown size={14} style={{ transform: showTaskReview ? "rotate(180deg)" : "none" }} />
                    </div>
                    {showTaskReview && (
                        <div style={S.draftBody}>
                            <div style={{ ...S.hint, marginBottom: "4px", color: "#065f46", fontSize: "10px" }}>Identificámos possíveis acionáveis neste email:</div>
                            <div style={{ display: "grid", gap: "6px" }}>
                                {extractedTasks.map((t, i) => (
                                    <div key={i} style={{ display: "flex", gap: "8px", alignItems: "flex-start" }}>
                                        <input
                                            type="checkbox"
                                            checked={t.completed}
                                            onChange={() => {
                                                const next = [...extractedTasks];
                                                next[i].completed = !next[i].completed;
                                                setExtractedTasks(next);
                                            }}
                                            style={{ marginTop: "3px", cursor: "pointer" }}
                                        />
                                        <div style={{ flex: 1, display: "flex", flexDirection: "column" }}>
                                            <input
                                                style={{ ...S.draftInput, background: "transparent", border: "none", padding: "0", fontWeight: 700, fontSize: "11px", color: t.completed ? "#94a3b8" : "#1e293b", textDecoration: t.completed ? "line-through" : "none" }}
                                                value={t.title}
                                                onChange={(e) => {
                                                    const next = [...extractedTasks];
                                                    next[i].title = e.target.value;
                                                    setExtractedTasks(next);
                                                }}
                                            />
                                            <div style={{ display: "flex", gap: "8px", fontSize: "9px", color: "#64748b", marginTop: "1px" }}>
                                                {t.dueDate && <span>📅 {t.dueDate}</span>}
                                                {t.owner && <span>👤 {t.owner}</span>}
                                            </div>
                                        </div>
                                    </div>
                                ))}
                            </div>
                            <div style={{ display: "flex", gap: "8px", marginTop: "8px", borderTop: "1px solid rgba(16, 185, 129, 0.1)", paddingTop: "8px" }}>
                                <button
                                    style={{ ...S.actionBtn, fontSize: "10px", color: "#059669" }}
                                    onClick={() => { void handleExtractTasksReview(); }}
                                >
                                    Atualizar
                                </button>
                                <button
                                    style={{ ...S.actionBtnPrimary, color: "#059669", display: "flex", alignItems: "center" }}
                                    onClick={async () => {
                                        const checklist = extractedTasks
                                            .filter(t => !t.completed)
                                            .map(t => `- [ ] ${t.title}${t.dueDate ? ` (${t.dueDate})` : ""}`)
                                            .join("\n");
                                        const ok = await copyTextWithFallback(`Lista de Tarefas:\n${checklist}`);
                                        setMsg(ok ? "Checklist copiada!" : "Nao foi possivel copiar automaticamente.");
                                    }}
                                >
                                    <Icons.Clipboard size={12} style={{ marginRight: "4px" }} />
                                    Copiar Checklist
                                </button>
                                <button
                                    style={{ ...S.actionBtn, fontSize: "10px", marginLeft: "auto" }}
                                    onClick={() => setShowTaskReview(false)}
                                >
                                    Recolher
                                </button>
                            </div>
                        </div>
                    )}
                </div>
            )}

            {/* Quick Knowledge Tools */}
            {bodyText && (bodyText.match(/\b\d{9}\b/) || bodyText.match(/\bPT50\b|\bIBAN\b/i)) && (
                <div style={{ ...S.briefingCard, background: "rgba(59, 130, 246, 0.05)", border: "1px dashed rgba(59, 130, 246, 0.3)", marginBottom: "8px" }}>
                    <div style={{ ...S.briefingHeader, color: "#1e40af" }}>
                        <Icons.Settings size={10} />
                        <span>Sugestão de Conhecimento</span>
                    </div>
                    <div style={{ ...S.briefingContent, fontSize: "10px", color: "#334155" }}>
                        Detetamos dados que podem ser úteis para futuras respostas (NIF/IBAN). Desejas guardar?
                    </div>
                    <div style={{ display: "flex", gap: "8px", marginTop: "4px" }}>
                        <button
                            style={{ ...S.actionBtnPrimary, fontSize: "10px" }}
                            onClick={() => {
                                // Simple heuristic: extract the first 9-digit number (NIF) or IBAN-like string
                                const nif = bodyText.match(/\b\d{9}\b/)?.[0];
                                const iban = bodyText.match(/\b(PT50\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{4}\s?\d{2})\b/i)?.[0] || bodyText.match(/IBAN:\s?(\S+)/i)?.[1];
                                if (nif) handleAddToKnowledge(`NIF: ${nif}`);
                                if (iban) handleAddToKnowledge(`IBAN: ${iban}`);
                            }}
                        >
                            Guardar Factos
                        </button>
                    </div>
                </div>
            )}

            <div style={S.inputCard}>
                <div style={S.actionToggleRow}>
                    <button
                        type="button"
                        style={selectedAction === "reply" ? S.actionToggleBtnOn : S.actionToggleBtn}
                        onClick={() => setGenerationAction("reply")}
                    >
                        Reply
                    </button>
                    <button
                        type="button"
                        style={selectedAction === "forward" ? S.actionToggleBtnOn : S.actionToggleBtn}
                        onClick={() => setGenerationAction("forward")}
                    >
                        Forward
                    </button>
                    {persistedEmailAttachments.length > 0 ? (
                        <div style={{ marginLeft: "auto", fontSize: "9px", color: "#64748b" }}>
                            {selectedAnalyzeFiles.length} analisar / {selectedForwardFiles.length} reenviar
                        </div>
                    ) : null}
                </div>
                {selectedAction === "reply" || selectedAction === "forward" ? (
                    <div style={S.replyTargetRow}>
                        <button
                            type="button"
                            style={S.replyTargetBtn}
                            onClick={() => void handlePickReplyTarget()}
                            title="Escolher um email guardado para usar como alvo desta resposta"
                        >
                            <Icons.MessageSquare size={11} />
                            Email alvo
                        </button>
                        {replyTargetEmail ? (
                            <div style={S.replyTargetSummary}>
                                <div style={{ display: "flex", alignItems: "center", gap: "6px", minWidth: 0 }}>
                                    <div style={S.replyTargetTitle}>{replyTargetEmail.subject || "(sem assunto)"}</div>
                                    <button type="button" style={S.replyTargetClear} onClick={() => setReplyTargetEmail(null)} title="Limpar email-alvo">×</button>
                                </div>
                                <div style={S.replyTargetMeta}>
                                    {replyTargetEmail.fromName || replyTargetEmail.fromEmail || "--"}
                                    {replyTargetEmail.receivedAtIso || replyTargetEmail.messageDateIso ? ` · ${formatDateLabel(replyTargetEmail.receivedAtIso || replyTargetEmail.messageDateIso || "")}` : ""}
                                </div>
                            </div>
                        ) : (
                            <div style={{ ...S.replyTargetSummary, borderColor: "rgba(148, 163, 184, 0.18)", background: "rgba(248,250,252,0.92)" }}>
                                <div style={{ ...S.replyTargetTitle, color: "#475569" }}>Sem email-alvo selecionado</div>
                                <div style={S.replyTargetMeta}>
                                    {selectedAction === "forward"
                                        ? "Abre a janela e escolhe o email guardado que queres usar como base deste reencaminhamento."
                                        : "Abre a janela e escolhe o email guardado a que queres responder usando este contexto."}
                                </div>
                            </div>
                        )}
                    </div>
                ) : null}
                {null}
                {persistedEmailAttachments.length > 0 && (
                    <div style={{ display: "flex", alignItems: "center", gap: "4px", marginBottom: "6px", padding: "2px 6px", background: "rgba(59, 130, 246, 0.05)", borderRadius: "4px", width: "fit-content" }}>
                        <Icons.Files size={10} color="var(--iccc-pill-active-bg)" />
                        <span style={{ fontSize: "10px", fontWeight: 700, color: "var(--iccc-pill-active-bg)" }}>
                            {persistedEmailAttachments.length} {persistedEmailAttachments.length === 1 ? "anexo disponivel" : "anexos disponiveis"}
                        </span>
                    </div>
                )}
                <textarea
                    style={S.textarea}
                    placeholder={selectedAction === "forward" ? "O que queres dizer no reencaminhamento?" : "O que queres escrever ou perguntar sobre este email?"}
                    value={prompt}
                    onChange={(e) => handlePromptChange(e.target.value)}
                    onKeyDown={(e) => handleKeyDown(e, selectedAction)}
                    onFocus={() => setDictationTarget("main")}
                />
                <div style={S.inputFooter}>
                    <div style={{ display: "flex", gap: "6px", alignItems: "center" }}>
                        <button
                            style={{
                                ...S.secondaryBtnPill,
                                width: "32px",
                                borderColor: isRecording ? "#ef4444" : "rgba(200, 210, 230, 0.6)",
                                color: isRecording ? "#ef4444" : "#172B4D",
                                background: isRecording ? "rgba(239, 68, 68, 0.1)" : "linear-gradient(180deg, rgba(255,255,255,0.95) 0%, rgba(220,228,245,0.85) 100%)",
                            }}
                            className="iccc-glossy-pill iccc-secondary-pill"
                            onClick={toggleRecording}
                            title="Ditado por voz"
                        >

                            <Icons.Microphone size={12} />
                        </button>
                        <button
                            className="iccc-glossy-pill iccc-secondary-pill"
                            style={S.secondaryBtnPill}
                            onClick={() => handleGenerate("summarize")}
                            disabled={isGenerating}
                            title={selectedAnalyzeFiles.length > 0 ? "Resumir email e anexos selecionados" : "Resumir este email"}
                            aria-label="Resumir"
                        >

                            <Icons.Receipt size={12} />
                        </button>
                        <button
                            className="iccc-glossy-pill iccc-secondary-pill"
                            style={S.secondaryBtnPill}
                            onClick={() => handleGenerate("tasks")}
                            disabled={isGenerating}
                            title="Extrair tarefas"
                            aria-label="Extrair"
                        >

                            <Icons.Check size={12} />
                        </button>
                        <button
                            className="iccc-glossy-pill iccc-secondary-pill"
                            style={S.secondaryBtnPill}
                            onClick={() => setActivePanel(activePanel === "files" ? null : "files")}
                            title="Selecionar anexos"
                        >

                            <Icons.Paperclip size={12} />
                        </button>
                    </div>
                    <button
                        className="iccc-glossy-pill iccc-primary-pill"
                        style={S.primaryBtnPill}
                        onClick={() => handleGenerate(selectedAction)}
                        disabled={isGenerating}
                    >

                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            {isGenerating ? "A GERAR..." : (selectedAction === "forward" ? "GERAR EMAIL" : "GERAR RESPOSTA")}
                            <Icons.Sparkles size={11} />
                        </div>
                    </button>
                </div>
            </div>

            <div style={S.refinerRow} ref={menuRef}>
                {/* Language Cascade */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{
                            ...S.secondaryBtnLink,
                            width: isNarrow ? "auto" : "64px",
                            minWidth: isNarrow ? "unset" : "64px",
                            justifyContent: "flex-start",
                            padding: isNarrow ? "0 6px" : "0 6px"
                        }}
                        onClick={() => setActivePanel(activePanel === "lang" ? null : "lang")}
                        title="Idioma"
                        aria-label="Idioma"
                    >
                        <div style={{
                            width: "16px",
                            height: "16px",
                            borderRadius: "50%",
                            overflow: "hidden",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            background: "rgba(0,0,0,0.05)",
                            fontSize: "11px",
                            lineHeight: 1,
                            flexShrink: 0,
                            boxShadow: "0 1px 2px rgba(0,0,0,0.12)"
                        }}>
                            <MiniFlag locale={selectedLocale || "auto"} />
                        </div>
                        {!isNarrow && (
                            <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>
                                {(selectedLocale || "auto") === "auto" ? "AUTO" : (selectedLocale || "auto").split("-")[0].toUpperCase()}
                            </span>
                        )}
                    </button>




                    {activeMenu === "lang" && (
                        <div id="FORCE_DOWN_MENU_LANG" style={S.cascadeMenu}>
                            {localeOptions.map((opt) => (
                                <button
                                    key={opt.value}
                                    className="iccc-glossy-pill iccc-secondary-pill"
                                    style={{ ...S.cascadeItem, background: "#ffffff", borderColor: "rgba(0,0,0,0.15)", width: "68px", minWidth: "68px" }}
                                    onClick={() => {
                                        setAiState({ locale: opt.value });
                                        setActiveMenu(null);
                                        if (output) handleGenerate("rewrite", output);
                                    }}
                                >
                                    <div style={{
                                        width: "16px",
                                        height: "16px",
                                        borderRadius: "50%",
                                        overflow: "hidden",
                                        display: "flex",
                                        alignItems: "center",
                                        justifyContent: "center",
                                        background: "rgba(0,0,0,0.05)",
                                        fontSize: "11px",
                                        lineHeight: 1,
                                        flexShrink: 0,
                                        boxShadow: "0 1px 2px rgba(0,0,0,0.12)"
                                    }}>
                                        <MiniFlag locale={opt.value} />
                                    </div>
                                    <span style={{ fontSize: "9px", fontWeight: 400 }}>
                                        {opt.value === "auto" ? "AUTO" : opt.value.split("-")[0].toUpperCase()}
                                    </span>
                                </button>



                            ))}
                        </div>
                    )}
                </div>

                <div style={{ width: "1px", height: "16px", background: "rgba(0,0,0,0.06)", margin: "0 2px" }}></div>

                {/* Presets Menu */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{
                            ...S.secondaryBtnLink,
                            width: isNarrow ? "auto" : "78px",
                            minWidth: isNarrow ? "unset" : "78px",
                            justifyContent: "flex-start",
                            padding: isNarrow ? "0 6px" : "0 8px"
                        }}
                        onClick={() => setActivePanel(activePanel === "presets" ? null : "presets")}
                        title="Modelos de Resposta Rápidos (MODS)"
                        aria-label="Modelos de Resposta Rápidos (MODS)"
                    >
                        <Icons.Settings size={11} style={{ opacity: 0.6 }} />
                        {!isNarrow && <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>MODS</span>}
                    </button>

                    {activeMenu === "presets" && (
                        <div style={{ ...S.cascadeMenu, width: "160px" }}>
                            <div style={{ padding: "4px 8px" }}>
                                <div style={{ ...S.chatInputWrapper, padding: "0 6px", background: "#fff", height: "24px" }}>
                                    <input
                                        style={{ ...S.chatInput, fontSize: "10px", padding: 0 }}
                                        placeholder="Procurar..."
                                        value={presetSearch}
                                        onChange={(e) => setPresetSearch(e.target.value)}
                                        autoFocus
                                    />
                                </div>
                            </div>

                            {(settings?.responsePresets || [])
                                .filter((p: any) =>
                                    !presetSearch ||
                                    p.name.toLowerCase().includes(presetSearch.toLowerCase()) ||
                                    p.prompt.toLowerCase().includes(presetSearch.toLowerCase())
                                )
                                .slice(0, 10) // Limit to avoid massive lists
                                .map((p: any) => (
                                    <button
                                        key={p.id}
                                        className="iccc-glossy-pill iccc-secondary-pill"
                                        style={S.cascadeItem}
                                        onClick={() => {
                                            setActiveMenu(null);
                                            setPresetSearch("");
                                            handleGenerate("reply", p.prompt);
                                        }}
                                    >
                                        <div style={{ width: "16px", display: "flex", justifyContent: "center" }}><Icons.ArrowRight size={10} /></div>
                                        <span style={{ fontWeight: 400, fontSize: "10px" }}>{p.name.toUpperCase()}</span>
                                    </button>
                                ))}

                            {settings?.responsePresets?.length === 0 && (
                                <div style={{ ...S.hint, padding: "10px", textAlign: "center", fontSize: "10px" }}>
                                    Cria modelos nas definições para acelerar respostas.
                                </div>
                            )}

                            {presetSearch && (settings?.responsePresets || []).filter((p: any) =>
                                p.name.toLowerCase().includes(presetSearch.toLowerCase()) ||
                                p.prompt.toLowerCase().includes(presetSearch.toLowerCase())
                            ).length === 0 && (
                                    <div style={{ ...S.hint, padding: "10px", textAlign: "center", fontSize: "10px" }}>
                                        Nenhum modelo encontrado.
                                    </div>
                                )}
                        </div>
                    )}
                </div>

                <div style={{ width: "1px", height: "16px", background: "rgba(0,0,0,0.06)", margin: "0 2px" }}></div>

                {/* Intents Menu (Smart Replies) */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{
                            ...S.secondaryBtnLink,
                            width: isNarrow ? "auto" : "72px",
                            minWidth: isNarrow ? "unset" : "72px",
                            justifyContent: "flex-start",
                            padding: isNarrow ? "0 6px" : "0 8px"
                        }}
                        onClick={toggleIntentsMenu}
                        disabled={isFetchingIntents}
                        title="Sugestões de Resposta da IA (DICAS)"
                        aria-label="Sugestões de Resposta da IA (DICAS)"
                    >
                        {isFetchingIntents ? (
                            <Icons.RotateCcw size={11} style={{ animation: "spin 1s linear infinite", opacity: 0.6 }} />
                        ) : (
                            <Icons.Activity size={11} style={{ opacity: 0.6 }} />
                        )}
                        {!isNarrow && <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>DICAS</span>}
                    </button>

                    {activeMenu === "intents" && (
                        <div style={{ ...S.cascadeMenu, width: "160px" }}>
                            <div style={{ display: "flex", justifyContent: "flex-end", padding: "6px 8px 0" }}>
                                <button
                                    style={{ ...S.actionBtn, fontSize: "9px" }}
                                    onClick={() => { void handleFetchIntents(true); }}
                                    disabled={isFetchingIntents}
                                >
                                    Atualizar
                                </button>
                            </div>
                            <div style={{ padding: "4px 8px" }}>
                                <div style={{ ...S.chatInputWrapper, padding: "0 6px", background: "#fff", height: "24px" }}>
                                    <input
                                        style={{ ...S.chatInput, fontSize: "10px", padding: 0 }}
                                        placeholder="Procurar..."
                                        value={intentSearch}
                                        onChange={(e) => setIntentSearch(e.target.value)}
                                        autoFocus
                                    />
                                </div>
                            </div>

                            {isFetchingIntents && (
                                <div style={{ ...S.hint, padding: "6px 10px", textAlign: "center", fontSize: "10px" }}>
                                    A gerar sugestoes...
                                </div>
                            )}

                            {aiState.smartReplies
                                .filter((i: string) => !intentSearch || i.toLowerCase().includes(intentSearch.toLowerCase()))
                                .map((intent: string, idx: number) => (
                                    <button
                                        key={idx}
                                        className="iccc-glossy-pill iccc-secondary-pill"
                                        style={S.cascadeItem}
                                        onClick={() => {
                                            setActiveMenu(null);
                                            setIntentSearch("");
                                            setPrompt(intent);
                                            handleGenerate("reply", intent);
                                        }}
                                    >
                                        <div style={{ width: "16px", display: "flex", justifyContent: "center" }}><Icons.Sparkles size={10} /></div>
                                        <span style={{ fontWeight: 400, fontSize: "10px", whiteSpace: "normal", overflowWrap: "anywhere", wordBreak: "break-word", maxWidth: "100%", flex: 1 }}>{intent.toUpperCase()}</span>
                                    </button>
                                ))}

                            {aiState.smartReplies.length === 0 && !isFetchingIntents && (
                                <div style={{ ...S.hint, padding: "10px", textAlign: "center", fontSize: "10px" }}>
                                    Nenhuma sugestão disponível.
                                </div>
                            )}
                        </div>
                    )}
                </div>

                <div style={{ width: "1px", height: "16px", background: "rgba(0,0,0,0.06)", margin: "0 2px" }}></div>

                {/* Contacts Menu */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{
                            ...S.secondaryBtnLink,
                            width: isNarrow ? "auto" : "72px",
                            minWidth: isNarrow ? "unset" : "72px",
                            justifyContent: "flex-start",
                            padding: isNarrow ? "0 6px" : "0 8px"
                        }}
                        onClick={toggleContactsMenu}
                        title="Contactos Sugeridos (LISTA)"
                        aria-label="Contactos Sugeridos (LISTA)"
                    >
                        {isExtractingContacts ? (
                            <Icons.RotateCcw size={11} style={{ animation: "spin 1s linear infinite", opacity: 0.6 }} />
                        ) : (
                            <Icons.User size={11} style={{ opacity: 0.6 }} />
                        )}
                        {!isNarrow && <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>LISTA</span>}
                    </button>

                    {activeMenu === "contacts" && (
                        <div style={{ ...S.cascadeMenu, width: "180px", left: "0" }}>
                            <div style={{ display: "flex", justifyContent: "flex-end", padding: "6px 8px 0" }}>
                                <button
                                    style={{ ...S.actionBtn, fontSize: "9px" }}
                                    onClick={() => { void handleExtractContacts(true); }}
                                    disabled={isExtractingContacts}
                                >
                                    Atualizar
                                </button>
                            </div>
                            <div style={{ padding: "4px 8px" }}>
                                <div style={{ ...S.chatInputWrapper, padding: "0 6px", background: "#fff", height: "24px" }}>
                                    <input
                                        style={{ ...S.chatInput, fontSize: "10px", padding: 0 }}
                                        placeholder="Procurar..."
                                        value={contactSearch}
                                        onChange={(e) => setContactSearch(e.target.value)}
                                        autoFocus
                                    />
                                </div>
                            </div>

                            {isExtractingContacts && (
                                <div style={{ ...S.hint, padding: "6px 10px", textAlign: "center", fontSize: "10px" }}>
                                    A detetar contactos...
                                </div>
                            )}

                            {/* Suggested from Email Context */}
                            {suggestedContacts.length > 0 && suggestedContacts
                                .filter(e => !contactSearch || e.toLowerCase().includes(contactSearch.toLowerCase()))
                                .map(email => (
                                    <button
                                        key={email}
                                        className="iccc-glossy-pill iccc-secondary-pill"
                                        style={S.cascadeItem}
                                        onClick={() => {
                                            setDraftTo((prev) => addUniqueEmail(prev, email));
                                            setActiveMenu(null);
                                        }}
                                    >
                                        <div style={{ width: "16px", display: "flex", justifyContent: "center" }}><Icons.AtSign size={10} /></div>
                                        <span style={{ fontSize: "9px", overflow: "hidden", textOverflow: "ellipsis" }}>{email}</span>
                                    </button>
                                ))}

                            {/* Aliases from Settings */}
                            {settings?.contactAliases?.length > 0 && settings.contactAliases
                                .filter((c: any) => !contactSearch || c.name.toLowerCase().includes(contactSearch.toLowerCase()) || c.email.toLowerCase().includes(contactSearch.toLowerCase()))
                                .map((c: any) => (
                                    <button
                                        key={c.id}
                                        className="iccc-glossy-pill iccc-secondary-pill"
                                        style={S.cascadeItem}
                                        onClick={() => {
                                            setDraftTo((prev) => addUniqueEmail(prev, c.email));
                                            setActiveMenu(null);
                                        }}
                                    >
                                        <div style={{ width: "16px", display: "flex", justifyContent: "center" }}><Icons.User size={10} /></div>
                                        <span style={{ fontWeight: 400 }}>{c.name.toUpperCase()}</span>
                                    </button>
                                ))}

                            {suggestedContacts.length === 0 && (settings?.contactAliases || []).length === 0 && (
                                <div style={{ ...S.hint, padding: "10px", textAlign: "center", fontSize: "10px" }}>
                                    Sem contactos detetados.
                                </div>
                            )}
                        </div>
                    )}
                </div>

                <div style={{ width: "1px", height: "16px", background: "rgba(0,0,0,0.06)", margin: "0 2px" }}></div>

                {/* Mode Cascade */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{ ...S.secondaryBtnLink, width: "64px", minWidth: "64px", justifyContent: "flex-start", padding: "0 6px" }}
                        onClick={() => setActivePanel(activePanel === "mode" ? null : "mode")}
                    >
                        <div style={{
                            width: "16px",
                            height: "16px",
                            borderRadius: "50%",
                            overflow: "hidden",
                            display: "flex",
                            alignItems: "center",
                            justifyContent: "center",
                            background: "rgba(0,0,0,0.05)",
                            lineHeight: 1,
                            flexShrink: 0,
                            boxShadow: "0 1px 2px rgba(0,0,0,0.12)"
                        }}>
                            {selectedCustomToneId ? <Icons.Sparkles size={11} /> : (toneRefiners.find(t => t.tone === baseTone)?.icon || <Icons.Settings size={11} />)}
                        </div>
                        <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>MODO</span>
                    </button>

                    {activeMenu === "mode" && (
                        <div id="FORCE_DOWN_MENU_MODE" style={S.cascadeMenu}>
                            {toneRefiners.map((r) => (
                                <button
                                    key={r.label}
                                    className="iccc-glossy-pill iccc-secondary-pill"
                                    style={{
                                        ...S.cascadeItem,
                                        background: aiState.tone === r.tone ? "rgba(37, 99, 235, 0.05)" : "#ffffff",
                                        color: aiState.tone === r.tone ? "#2563eb" : "#172B4D",
                                        borderColor: aiState.tone === r.tone ? "#2563eb" : "rgba(0,0,0,0.1)"
                                    }}
                                    onClick={() => {
                                        setAiState({ tone: r.tone });
                                        setActiveMenu(null);
                                        if (output) handleGenerate("rewrite", output);
                                    }}
                                >
                                    <div style={{ width: "16px", display: "flex", justifyContent: "center" }}>{r.icon}</div>
                                    <span style={{ fontWeight: 400 }}>{r.label.toUpperCase()}</span>
                                </button>


                            ))}
                        </div>
                    )}
                </div>
            </div>

            {renderQuickPanel()}

            {generationError ? (
                <div
                    style={{
                        border: "1px solid rgba(239, 68, 68, 0.24)",
                        background: "rgba(254, 242, 242, 0.96)",
                        color: "#991b1b",
                        borderRadius: "10px",
                        padding: "8px 10px",
                        fontSize: "11px",
                        lineHeight: 1.35,
                        marginTop: "4px",
                        marginBottom: "4px",
                    }}
                >
                    <strong>Erro na geração:</strong> {generationError}
                </div>
            ) : null}

            {
                (output || isGenerating || aiState.history.length > 0 || history.length > 0) && (
                    <div style={S.outputCard}>
                        <div style={S.outputHeader}>
                            <div style={{ display: "flex", alignItems: "center", gap: "6px" }} title="Sugestões da IA">
                                <Icons.Sparkles size={15} style={{ opacity: 0.6 }} />
                                {isGenerating && <div style={S.typingDots}><span>.</span><span>.</span><span>.</span></div>}
                            </div>
                            <div style={{ display: "flex", gap: "8px", alignItems: "center" }}>
                                {history.length > 0 && (
                                    <div style={{ position: "relative" }}>
                                        <button
                                            style={{
                                                ...S.actionBtn,
                                                color: "#2563eb",
                                                display: "flex",
                                                alignItems: "center",
                                                background: !output ? "rgba(37, 99, 235, 0.08)" : "none",
                                                padding: !output ? "4px 8px" : "0",
                                                borderRadius: "4px"
                                            }}
                                            onClick={() => setHistoryExpanded((prev) => !prev)}
                                            title="Histórico"
                                            aria-label="Histórico"
                                        >
                                            <Icons.RotateCcw size={15} />
                                        </button>
                                        {String(activeMenu || "") === "rollback_hidden" && (
                                            <div style={{ ...S.cascadeMenu, width: "220px", right: 0, left: "auto", top: "24px" }}>
                                                {history.map((h, i) => (
                                                    <button
                                                        key={h.id}
                                                        className="iccc-glossy-pill iccc-secondary-pill"
                                                        style={{ ...S.cascadeItem, height: "auto", padding: "6px 10px", flexDirection: "column", alignItems: "flex-start", gap: "2px" }}
                                                        onClick={() => {
                                                            applyHistoryEntry(h);
                                                            setMsg("Versão anterior restaurada.");
                                                        }}
                                                    >
                                                        <div style={{ display: "flex", alignItems: "center", gap: "4px", width: "100%" }}>
                                                            <Icons.Clock size={10} style={{ opacity: 0.5 }} />
                                                            <span style={{ fontSize: "8px", color: "#64748b" }}>{new Date(h.ts).toLocaleString()}</span>
                                                            <span style={{ marginLeft: "auto", fontSize: "8px", color: "#475569" }}>{i === 0 ? "Atual" : "Anterior"}</span>
                                                        </div>
                                                        <div style={{ fontSize: "10px", color: "#1e293b", fontWeight: 700, whiteSpace: "normal", textAlign: "left" }}>
                                                            {h.prompt.length > 40 ? h.prompt.substring(0, 40) + "..." : h.prompt || "(Sem instrução)"}
                                                        </div>
                                                    </button>
                                                ))}
                                            </div>
                                        )}
                                    </div>
                                )}
                                <button style={S.actionBtn} onClick={handleResetConversation} title="Eliminar" aria-label="Eliminar">
                                    <Icons.Trash size={15} />
                                </button>
                                <button style={S.actionBtn} onClick={handleExport} title="Download" aria-label="Download">
                                    <Icons.Download size={15} />
                                </button>
                                <button
                                    style={S.actionBtn}
                                    onClick={async () => {
                                        const ok = await copyTextWithFallback(output);
                                        setMsg(ok ? "Texto copiado." : "Nao foi possivel copiar automaticamente.");
                                    }}
                                    title="Copiar"
                                    aria-label="Copiar"
                                >
                                    <Icons.Clipboard size={15} />
                                </button>
                                <button
                                    style={S.actionBtnPrimary}
                                    onClick={async () => {
                                        const settings = await getSettings();
                                        const mLinks = settings.meetingLinks;
                                        const mLink = mLinks?.teams || mLinks?.zoom || mLinks?.meet || "";
                                        const finalBody = mLink ? `${output}\n\n---\nLink da Reunião: ${mLink}` : output;

                                        await displayNewMeetingForm({
                                            subject: ctx.subject ? `Re: ${ctx.subject}` : "Reunião",
                                            body: finalBody,
                                            requiredAttendees: ctx.fromEmail ? [ctx.fromEmail] : []
                                        });
                                    }}
                                    title="Agendar"
                                    aria-label="Agendar"
                                >
                                    <Icons.Calendar size={15} />
                                </button>
                                <button
                                    style={S.actionBtnPrimary}
                                    onClick={async () => {
                                        try {
                                            await handleInsert();
                                        } catch (e: any) {
                                            alert("Erro crítico no botão: " + e.message);
                                        }
                                    }}
                                    title="Inserir"
                                    aria-label="Inserir"
                                >
                                    <Icons.Send size={15} />
                                </button>
                            </div>
                        </div>

                        {historyExpanded && (
                            <div style={S.quickPanel}>
                                <div style={S.quickPanelHeader}>
                                    <span>Histórico</span>
                                    <button type="button" style={{ ...S.actionBtn, fontSize: "10px" }} onClick={() => setHistoryExpanded(false)}>
                                        Fechar
                                    </button>
                                </div>
                                <div style={S.quickPanelBody}>
                                    {history.map((entry, i) => (
                                        <button
                                            key={entry.id}
                                            type="button"
                                            style={{ ...S.quickPanelItemBtn, flexDirection: "column", alignItems: "flex-start" }}
                                            onClick={() => {
                                                applyHistoryEntry(entry);
                                                setHistoryExpanded(false);
                                                setMsg("Versão anterior restaurada.");
                                            }}
                                        >
                                            <div style={{ display: "flex", alignItems: "center", gap: "6px", width: "100%" }}>
                                                <Icons.Clock size={10} style={{ opacity: 0.6 }} />
                                                <span style={{ fontSize: "9px", color: "#64748b" }}>{new Date(entry.ts).toLocaleString()}</span>
                                                <span style={{ marginLeft: "auto", fontSize: "9px", color: "#475569" }}>{i === 0 ? "Atual" : String(entry.action || "reply").toUpperCase()}</span>
                                            </div>
                                            <div style={{ fontSize: "11px", fontWeight: 700, color: "#172B4D", whiteSpace: "normal" }}>
                                                {entry.prompt.length > 64 ? `${entry.prompt.slice(0, 64)}...` : entry.prompt || "(Sem instrução)"}
                                            </div>
                                            <div style={{ fontSize: "9px", color: "#64748b", whiteSpace: "normal" }}>
                                                {entry.output.length > 96 ? `${htmlToPlainText(entry.output).slice(0, 96)}...` : htmlToPlainText(entry.output)}
                                            </div>
                                        </button>
                                    ))}
                                    {history.length === 1 ? (
                                        <div style={{ ...S.quickPanelItem, color: "#64748b" }}>Existe apenas a versÃ£o atual.</div>
                                    ) : null}
                                </div>
                            </div>
                        )}

                        {showDraftPreview && (
                            <div style={S.draftCard}>
                                <div style={S.draftHeader} onClick={() => setDraftDetailsExpanded((prev) => !prev)}>
                                    <Icons.Settings size={12} />
                                    <span>Detalhes do Rascunho</span>
                                    <Icons.ArrowDown size={14} style={{ marginLeft: "auto", transform: draftDetailsExpanded ? "rotate(180deg)" : "none" }} />
                                </div>
                                {draftDetailsExpanded && (
                                <div style={S.draftBody}>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>Para:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftTo.join("; ")}
                                            onChange={(e) => setDraftTo(normalizeEmailListInput(e.target.value))}
                                            placeholder="exemplo@mail.com; ..."
                                            title="Destinatários principais"
                                        />
                                    </div>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>CC:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftCc.join("; ")}
                                            onChange={(e) => setDraftCc(normalizeEmailListInput(e.target.value))}
                                            placeholder="cc@mail.com; ..."
                                            title="Destinatários em cópia"
                                        />
                                    </div>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>Assunto:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftSubject}
                                            onChange={(e) => setDraftSubject(e.target.value)}
                                            placeholder="Assunto do email"
                                            title="Assunto"
                                        />
                                    </div>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>Bcc:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftBcc.join("; ")}
                                            onChange={(e) => setDraftBcc(normalizeEmailListInput(e.target.value))}
                                            placeholder="bcc@mail.com; ..."
                                            title="Destinatarios em copia oculta"
                                        />
                                    </div>
                                    {draftRecipientSuggestions.length > 0 ? (
                                        <div style={{ display: "grid", gap: "4px", marginTop: "2px" }}>
                                            <div style={{ fontSize: "9px", color: "#64748b", fontWeight: 700, textTransform: "uppercase" }}>
                                                Sugestoes
                                            </div>
                                            {draftRecipientSuggestions.map((email) => (
                                                <div key={email} style={S.quickPanelItem}>
                                                    <span style={{ flex: 1, minWidth: 0, overflow: "hidden", textOverflow: "ellipsis", fontSize: "10px" }}>{email}</span>
                                                    <div style={{ display: "flex", gap: "4px" }}>
                                                        <button type="button" style={draftTo.includes(email) ? S.purposeChipOn : S.purposeChip} onClick={() => addDraftRecipient("to", email)}>Para</button>
                                                        <button type="button" style={draftCc.includes(email) ? S.purposeChipOn : S.purposeChip} onClick={() => addDraftRecipient("cc", email)}>Cc</button>
                                                        <button type="button" style={draftBcc.includes(email) ? S.purposeChipOn : S.purposeChip} onClick={() => addDraftRecipient("bcc", email)}>Bcc</button>
                                                    </div>
                                                </div>
                                            ))}
                                        </div>
                                    ) : null}

                                </div>
                                )}
                            </div>
                        )}

                        {/* Iterative Result Area */}
                        <div style={{ position: "relative" }}>
                            <div style={S.outputText} dangerouslySetInnerHTML={{ __html: output }} />
                            {isGenerating && !output && (
                                <div style={{ padding: "20px 0", color: "var(--iccc-text-muted)", fontStyle: "italic" }}>
                                    A pensar...
                                </div>
                            )}
                        </div>

                        {/* Quick Refinement Input */}
                        <div style={S.chatInputWrapper}>
                            <textarea
                                style={{ ...S.chatInput, height: "unset", minHeight: "32px", maxHeight: "120px", paddingTop: "8px", resize: "none" }}
                                placeholder="Refinar resposta (ex: faz mais curto)..."
                                value={refineInput}
                                onChange={(e) => setRefineInput(e.target.value)}
                                onKeyDown={(e) => handleKeyDown(e, "refine")}
                                onFocus={() => setDictationTarget("refine")}
                                rows={1}
                            />
                            <button
                                style={{
                                    ...S.chatSendBtn,
                                    width: "24px",
                                    height: "22px",
                                    background: isRecording ? "rgba(239, 68, 68, 0.1)" : "transparent",
                                    color: isRecording ? "#ef4444" : "var(--iccc-text-muted)",
                                    border: isRecording ? "1px solid #ef4444" : "none"
                                }}
                                onClick={toggleRecording}
                                title="Ditado por voz"
                            >
                                <Icons.Microphone size={12} />
                            </button>
                            <button
                                disabled={isGenerating || !refineInput}
                                onClick={() => {
                                    handleGenerate("refine");
                                    setRefineInput("");
                                }}
                                style={{
                                    ...S.chatSendBtn,
                                    opacity: !refineInput || isGenerating ? 0.5 : 1
                                }}
                            >
                                <Icons.Sparkles size={14} />
                            </button>
                        </div>
                    </div>
                )
            }
        </div>
    );
};

