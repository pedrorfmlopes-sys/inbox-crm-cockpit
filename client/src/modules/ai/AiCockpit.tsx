import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { aiGenerate, type AiAction, type AiTone, type AiLocale } from "@/ai/aiClient";
import { insertTextToBody, isComposeMode, displayReplyForm, displayForwardForm, displayNewMeetingForm, setRecipients, setSubject } from "@/office";
import { getSettings } from "@/settings";
import { logLearningInteraction } from "@/api";
import * as Icons from "@/ui/icons";

// --- PERSISTENCE HELPERS ---
const AI_HISTORY_KEY = "icc.ai_history.v1";
const HISTORY_KEEP_MS = 5 * 24 * 60 * 60 * 1000; // 5 days

type HistoryEntry = {
    id: string;
    emailKey: string;
    ts: number;
    output: string;
    prompt: string;
    action: string;
};

function getEmailKey(ctx: any) {
    return ctx.conversationId || ctx.internetMessageId || "global";
}

function loadHistory(): HistoryEntry[] {
    try {
        const raw = localStorage.getItem(AI_HISTORY_KEY);
        if (!raw) return [];
        const arr = JSON.parse(raw);
        const now = Date.now();
        return arr.filter((h: any) => now - h.ts < HISTORY_KEEP_MS);
    } catch { return []; }
}

function saveHistory(entries: HistoryEntry[]) {
    try {
        const now = Date.now();
        const pruned = entries.filter(h => now - h.ts < HISTORY_KEEP_MS).slice(-100);
        localStorage.setItem(AI_HISTORY_KEY, JSON.stringify(pruned));
    } catch { }
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

function parseExtractedTasks(rawText: string): Array<{ title: string; dueDate?: string; owner?: string }> {
    const trimmed = String(rawText || "").trim();
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
            const list = Array.isArray(parsed)
                ? parsed
                : Array.isArray(parsed?.tasks)
                    ? parsed.tasks
                    : [];
            const normalized = list
                .map((task: any) => ({
                    title: String(task?.title || task?.task || task?.name || "").trim(),
                    dueDate: String(task?.dueDate || task?.due || "").trim() || undefined,
                    owner: String(task?.owner || task?.assignee || "").trim() || undefined,
                }))
                .filter((task: any) => task.title);
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
    const { ctx, bodyText, bodyHtml, setMsg, aiState, setAiState, files, addFile, clearFiles, settings } = useCockpit() as any;
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

    // Draft Preview State
    const [draftTo, setDraftTo] = useState<string[]>([]);
    const [draftCc, setDraftCc] = useState<string[]>([]);
    const [draftSubject, setDraftSubject] = useState("");
    const [suggestedContacts, setSuggestedContacts] = useState<string[]>([]);
    const [showDraftPreview, setShowDraftPreview] = useState(false);
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

    // History / Rollback
    const [history, setHistory] = useState<HistoryEntry[]>([]);
    const emailKey = getEmailKey(ctx);

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
        setHistory(loadHistory().filter(h => h.emailKey === emailKey));
    }, [ctx.conversationId, emailKey]);

    useEffect(() => {
        setBriefing(null);
        setBriefingExpanded(false);
        setSuggestedContacts([]);
        setExtractedTasks([]);
        setShowTaskReview(false);
        setContactSearch("");
        setIntentSearch("");
    }, [emailKey]);

    useEffect(() => {
        if (aiManualOnly || !ctx.conversationId || ctx.isCompose) return;

        void handleFetchBriefing(false);
        void handleFetchIntents(false);
        if (bodyText) void handleExtractContacts(false);
        if ((bodyText || "").length >= 50) void handleExtractTasksReview();
    }, [aiManualOnly, emailKey, ctx.conversationId, ctx.isCompose, bodyText]);

    async function handleExtractTasksReview() {
        const effectiveBodyText = String(bodyText || "").trim() || htmlToPlainText(bodyHtml || "");
        if (!ctx.conversationId || ctx.isCompose || isExtractingTasks) return;
        if (!effectiveBodyText) {
            setMsg("O corpo deste email ainda não ficou disponível no Outlook. Tenta novamente dentro de 1-2 segundos.");
            return;
        }
        if (effectiveBodyText.length < 50) {
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
                    bodyText: effectiveBodyText,
                } as any
            });
            if (!res.ok) {
                setExtractedTasks([]);
                setShowTaskReview(false);
                setMsg(res.error || "Erro ao extrair tarefas.");
                return;
            }

            const tasks = parseExtractedTasks(res.text || "");
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
                    bodyText: bodyText || "",
                    bodyScope: nextSettings.bodyScope || "main"
                },
                briefing: briefing,
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
            const res = await aiGenerateBriefing(bodyText || "", [], {}, ctx.conversationId);
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
                    bodyText,
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
        const nextOpen = activeMenu !== "intents";
        setActiveMenu(nextOpen ? "intents" : null);
        if (nextOpen && !aiState.smartReplies.length) {
            void handleFetchIntents(true);
        }
    }

    function toggleContactsMenu() {
        const nextOpen = activeMenu !== "contacts";
        setActiveMenu(nextOpen ? "contacts" : null);
        if (nextOpen && !suggestedContacts.length) {
            void handleExtractContacts(true);
        }
    }

    // Sync draft defaults from context OR persistent aiState
    useEffect(() => {
        // If we have AI-suggested metadata, use it
        if (aiState.suggestedSubject || (aiState.suggestedTo && aiState.suggestedTo.length > 0)) {
            setDraftTo(aiState.suggestedTo || []);
            setDraftCc(aiState.suggestedCc || []);
            setDraftSubject(aiState.suggestedSubject || "");
        } else {
            // Fallback to email context defaults
            setDraftTo((ctx.toRecipients || []).map((r: any) => r.email));
            setDraftCc((ctx.ccRecipients || []).map((r: any) => r.email));
            setDraftSubject(ctx.subject || "");
        }
    }, [ctx, aiState.suggestedSubject, aiState.suggestedTo]);

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
        recognition.lang = aiState.locale === "auto" ? "pt-PT" : aiState.locale as any;
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
            const { getAttachments } = await import("@/office");
            const atts = await getAttachments();
            if (atts.length === 0) {
                setMsg("Nenhum anexo encontrado neste email.");
            } else {
                let counts = 0;
                for (const att of atts) {
                    const ctxFiles = files || [];
                    if (!ctxFiles.find((f: any) => f.name === att.name)) {
                        addFile({
                            name: att.name,
                            type: att.contentType,
                            content: att.content
                        });
                        counts++;
                    }
                }
                if (counts > 0) setMsg(`${counts} anexos importados!`);
                else setMsg("Anexos já estavam carregados.");
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

    async function handleGenerate(action: AiAction = "reply", extraPrompt?: string) {
        if (isGenerating) return;
        setIsGenerating(true);
        // If we are starting a NEW task (not refining), clear previous history
        const isRefining = action === "rewrite" || action === "refine";
        const finalPrompt = extraPrompt || (action === "refine" ? refineInput : prompt);

        if (!isRefining && aiState.history.length > 0) {
            setAiState({ history: [] });
        }

        setOutput("");
        setDebugLog("");

        try {
            const settings = await getSettings();
            const res = await aiGenerate({
                action,
                mode: "quality",
                tone: aiState.tone || settings.tone || "neutro",
                locale: aiState.locale || settings.replyLanguage || "pt-PT",
                inputText: finalPrompt,
                files: files || [],
                briefing: briefing, // Pass the thread summary for isolation
                email: {
                    subject: ctx.subject || "",
                    from: ctx.fromEmail || "",
                    to: (ctx.toRecipients || []).map((r: any) => r.email),
                    cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                    bodyText: bodyText || "",
                    bodyScope: settings.bodyScope || "main"
                },
                persona: {
                    userRole: settings.userRole,
                    styleContext: settings.styleContext,
                    styleExamples: settings.styleExamples,
                },
                history: isRefining ? aiState.history : [],
                knowledge: settings.aiKnowledge || [],
                contactAliases: settings.contactAliases || [],
                // For refine: send the current editor content as explicit draft
                draftText: action === "refine" ? (output || aiState.output || "") : undefined,
            }); //inputText is already extraPrompt || prompt

            if (res.ok) {
                setAiState({
                    action,
                    output: res.text,
                    suggestedTo: res.suggestedRecipients?.to || [],
                    suggestedCc: res.suggestedRecipients?.cc || [],
                    suggestedSubject: res.suggestedSubject || ""
                });
                let fullText = res.text;
                let current = "";
                const words = fullText.split(" ");
                for (let i = 0; i < words.length; i++) {
                    current += words[i] + " ";
                    setOutput(current);
                    await new Promise((resolve) => setTimeout(resolve, 20));
                }

                const newHistory = [
                    ...(isRefining ? aiState.history : []),
                    { role: "user" as const, content: finalPrompt },
                    { role: "assistant" as const, content: fullText }
                ].slice(-4);
                setAiState({ output: fullText, history: newHistory });
                setPrompt("");
                setShowDraftPreview(true);

                // Persist to Local History
                const entry: HistoryEntry = {
                    id: Math.random().toString(36).substring(7),
                    emailKey,
                    ts: Date.now(),
                    output: fullText,
                    prompt: finalPrompt,
                    action
                };
                const fullHist = [entry, ...loadHistory()];
                saveHistory(fullHist);
                setHistory(fullHist.filter(h => h.emailKey === emailKey));
            } else {
                setMsg(res.error);
            }
        } catch (e: any) {
            setMsg(e.message);
        } finally {
            setIsGenerating(false);
        }
    }

    async function handleInsert() {
        console.log("[AiCockpit] handleInsert called");
        setDebugLog("Botão clicado. A verificar modo...");
        try {
            const isCompose = await isComposeMode();
            console.log("[AiCockpit] isComposeMode:", isCompose);
            setDebugLog(`Modo Edição: ${isCompose}`);

            if (isCompose) {
                setDebugLog("A atualizar rascunho...");

                // Sync metadata first
                await setRecipients("to", draftTo);
                await setRecipients("cc", draftCc);
                await setSubject(draftSubject);

                // Insert body
                await insertTextToBody(output);

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
            const effectiveAction = aiState.action || "reply";

            if (effectiveAction === "forward") {
                await displayForwardForm(output);
            } else {
                // Default to Reply (including for refine, rewrite, etc.)
                await displayReplyForm(output);
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

    const [activeMenu, setActiveMenu] = useState<"lang" | "mode" | "presets" | "intents" | "contacts" | null>(null);
    const menuRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (menuRef.current && !menuRef.current.contains(e.target as Node)) setActiveMenu(null);
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

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
            suggestedSubject: ""
        });
        setPrompt("");
        setOutput("");
        setBriefing(null);
        setExtractedTasks([]);
        setShowTaskReview(false);
        setDraftTo([]);
        setDraftCc([]);
        setDraftSubject("");
        setSuggestedContacts([]);
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
    };


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
                                    onClick={() => {
                                        const checklist = extractedTasks
                                            .filter(t => !t.completed)
                                            .map(t => `- [ ] ${t.title}${t.dueDate ? ` (${t.dueDate})` : ""}`)
                                            .join("\n");
                                        navigator.clipboard.writeText(`Lista de Tarefas:\n${checklist}`);
                                        setMsg("Checklist copiada!");
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
                {files && files.length > 0 && (
                    <div style={{ display: "flex", alignItems: "center", gap: "4px", marginBottom: "6px", padding: "2px 6px", background: "rgba(59, 130, 246, 0.05)", borderRadius: "4px", width: "fit-content" }}>
                        <Icons.Files size={10} color="var(--iccc-pill-active-bg)" />
                        <span style={{ fontSize: "10px", fontWeight: 700, color: "var(--iccc-pill-active-bg)" }}>
                            {files.length} {files.length === 1 ? "anexo pronto" : "anexos prontos"}
                        </span>
                    </div>
                )}
                <textarea
                    style={S.textarea}
                    placeholder="O que queres escrever ou perguntar sobre este email?"
                    value={prompt}
                    onChange={(e) => handlePromptChange(e.target.value)}
                    onKeyDown={(e) => handleKeyDown(e, "reply")}
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
                            title={files.length > 0 ? "Resumir email e anexos identificados" : "Resumir este email"}
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
                            onClick={() => handleGenerate("forward")}
                            disabled={isGenerating}
                            title="Reenviar (Rascunho)"
                        >

                            <Icons.Send size={12} />
                        </button>
                        <button
                            className="iccc-glossy-pill iccc-secondary-pill"
                            style={S.secondaryBtnPill}
                            onClick={handleImportAttachments}
                            disabled={isImporting}
                            title="Importar anexos deste email"
                        >

                            {isImporting ? <Icons.RotateCcw size={12} style={{ animation: "spin 1s linear infinite" }} /> : <Icons.Paperclip size={12} />}
                        </button>
                    </div>
                    <button
                        className="iccc-glossy-pill iccc-primary-pill"
                        style={S.primaryBtnPill}
                        onClick={() => handleGenerate("reply")}
                        disabled={isGenerating}
                    >

                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            {isGenerating ? "A GERAR..." : "GERAR RESPOSTA"}
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
                        onClick={() => setActiveMenu(activeMenu === "lang" ? null : "lang")}
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
                            <MiniFlag locale={aiState.locale || "auto"} />
                        </div>
                        {!isNarrow && (
                            <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 400 }}>
                                {(aiState.locale || "auto").split("-")[0].toUpperCase()}
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
                        onClick={() => setActiveMenu(activeMenu === "presets" ? null : "presets")}
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
                                            if (!draftTo.includes(email)) setDraftTo([...draftTo, email]);
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
                                            if (!draftTo.includes(c.email)) setDraftTo([...draftTo, c.email]);
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
                        onClick={() => setActiveMenu(activeMenu === "mode" ? null : "mode")}
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
                            {toneRefiners.find(t => t.tone === aiState.tone)?.icon || <Icons.Settings size={11} />}
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
                                            onClick={() => setActiveMenu(activeMenu === "rollback" as any ? null : "rollback" as any)}
                                            title="Histórico"
                                            aria-label="Histórico"
                                        >
                                            <Icons.RotateCcw size={15} />
                                        </button>
                                        {activeMenu === "rollback" as any && (
                                            <div style={{ ...S.cascadeMenu, width: "220px", right: 0, left: "auto", top: "24px" }}>
                                                {history.slice(1).map((h, _i) => (
                                                    <button
                                                        key={h.id}
                                                        className="iccc-glossy-pill iccc-secondary-pill"
                                                        style={{ ...S.cascadeItem, height: "auto", padding: "6px 10px", flexDirection: "column", alignItems: "flex-start", gap: "2px" }}
                                                        onClick={() => {
                                                            setOutput(h.output);
                                                            setActiveMenu(null);
                                                            setMsg("Versão anterior restaurada.");
                                                        }}
                                                    >
                                                        <div style={{ display: "flex", alignItems: "center", gap: "4px", width: "100%" }}>
                                                            <Icons.Clock size={10} style={{ opacity: 0.5 }} />
                                                            <span style={{ fontSize: "8px", color: "#64748b" }}>{new Date(h.ts).toLocaleString()}</span>
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
                                <button style={S.actionBtn} onClick={() => navigator.clipboard.writeText(output)} title="Copiar" aria-label="Copiar">
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

                        {showDraftPreview && (
                            <div style={S.draftCard}>
                                <div style={S.draftHeader} onClick={() => setShowDraftPreview(!showDraftPreview)}>
                                    <Icons.Settings size={12} />
                                    <span>Detalhes do Rascunho</span>
                                    <Icons.ArrowDown size={14} style={{ marginLeft: "auto", transform: showDraftPreview ? "rotate(180deg)" : "none" }} />
                                </div>
                                <div style={S.draftBody}>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>Para:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftTo.join("; ")}
                                            onChange={(e) => setDraftTo(e.target.value.split(";").map(v => v.trim()))}
                                            placeholder="exemplo@mail.com; ..."
                                            title="Destinatários principais"
                                        />
                                    </div>
                                    <div style={S.draftRow}>
                                        <label style={S.draftLabel}>CC:</label>
                                        <input
                                            style={S.draftInput}
                                            value={draftCc.join("; ")}
                                            onChange={(e) => setDraftCc(e.target.value.split(";").map(v => v.trim()))}
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

                                </div>
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


