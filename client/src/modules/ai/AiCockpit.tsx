import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { aiGenerate, type AiAction, type AiTone, type AiLocale } from "@/ai/aiClient";
import { insertTextToBody, isComposeMode, displayReplyForm, displayForwardForm, displayNewMeetingForm, setRecipients, setSubject } from "@/office";
import { getSettings } from "@/settings";
import * as Icons from "@/ui/icons";

export const AiCockpit: React.FC = () => {
    const { ctx, bodyText, setMsg, aiState, setAiState, files, addFile, removeFile, clearFiles, settings } = useCockpit() as any;

    // Local state for immediate typing feel
    // Initialized from context, but NOT synced on every render to avoid loops
    const [prompt, setPrompt] = useState(aiState.prompt);
    const [output, setOutput] = useState(aiState.output);
    const [briefing, setBriefing] = useState<string | null>(null);
    const [isGenerating, setIsGenerating] = useState(false);
    const [isFetchingIntents, setIsFetchingIntents] = useState(false);
    const [isFetchingBriefing, setIsFetchingBriefing] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
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
    const recognitionRef = useRef<any>(null);

    // Presets Search State
    const [presetSearch, setPresetSearch] = useState("");

    // CRITICAL: Only sync local state from context when the conversation (email) changes.
    useEffect(() => {
        setPrompt(aiState.prompt);
        setOutput(aiState.output);
        setDebugLog(""); // Clear debug log on switch
    }, [ctx.conversationId]);

    // Automated Task Extraction in Read Mode (with Persistence)
    useEffect(() => {
        // Only clear tasks if the conversation changed
        // We don't clear when entering Compose if we already have tasks for this email
        if (!ctx.conversationId || !bodyText) {
            setExtractedTasks([]);
            setShowTaskReview(false);
            return;
        }

        // If we are already in Compose, we don't trigger a new extraction automatically
        // but we keep what was found in Read mode.
        if (ctx.isCompose) return;

        // If we already have tasks and the review is shown, don't re-extract
        if (extractedTasks.length > 0 && showTaskReview) return;

        // Smart Filter: Skip very short emails (likely unrelated to tasks)
        if (bodyText.length < 50) return;

        const extractTasks = async () => {
            setIsExtractingTasks(true);
            try {
                const res = await aiGenerate({
                    action: "extract_tasks_json" as any,
                    mode: "fast",
                    locale: "pt-PT",
                    tone: "neutro",
                    email: {
                        subject: ctx.subject || "",
                        from: ctx.fromEmail || "",
                        to: (ctx.toRecipients || []).map((r: any) => r.email),
                        cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                        bodyText: bodyText || "",
                    } as any
                });
                if (res.ok) {
                    try {
                        const json = JSON.parse(res.text.trim());
                        if (Array.isArray(json) && json.length > 0) {
                            setExtractedTasks(json.map(t => ({ ...t, completed: false })));
                            setShowTaskReview(true);
                        }
                    } catch (e) {
                        console.error("Failed to parse tasks JSON:", res.text);
                    }
                }
            } catch (err) {
                console.error("Erro ao extrair tarefas:", err);
            } finally {
                setIsExtractingTasks(false);
            }
        };

        const timer = setTimeout(extractTasks, 1500); // Delay to ensure context is ready
        return () => clearTimeout(timer);
    }, [ctx.conversationId, ctx.isCompose, bodyText]);

    // Automated Intent Proposals
    useEffect(() => {
        if (!ctx.conversationId || ctx.isCompose) {
            setAiState({ smartReplies: [] });
            return;
        }

        const fetchIntents = async () => {
            setIsFetchingIntents(true);
            try {
                const settings = await getSettings();
                const res = await aiGenerate({
                    action: "intent_proposals", // This was explicitly "intent_proposals"
                    mode: "fast",
                    locale: (settings.replyLanguage || "pt-PT") as any, // This was from settings
                    tone: settings.tone || "neutro", // This was from settings
                    email: {
                        subject: ctx.subject || "",
                        from: ctx.fromEmail || "",
                        to: (ctx.toRecipients || []).map((r: any) => r.email),
                        cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                        bodyText: bodyText || "",
                        bodyScope: settings.bodyScope || "main" // Added this line
                    },
                    // inputText, knowledge, history, files are not used for intent_proposals
                    // and would require 'action', 'extraPrompt', 'prompt', 'isRefining' to be defined.
                    // Keeping the original structure for intent_proposals and adding briefing.
                    briefing: briefing, // Pass the briefing for isolation
                    persona: {
                        userRole: settings.userRole,
                        styleContext: settings.styleContext,
                        styleExamples: settings.styleExamples,
                    }
                });
                if (res.ok) {
                    const intents = res.text.split(";").map(i => i.trim()).filter(Boolean);
                    setAiState({ smartReplies: intents });
                }
            } catch (err) {
                console.error("Erro ao obter intenções:", err);
            } finally {
                setIsFetchingIntents(false);
            }
        };

        fetchIntents();
    }, [ctx.conversationId, ctx.isCompose]);

    // Automated Briefing on Conversation Change
    useEffect(() => {
        if (!ctx.conversationId || ctx.isCompose) return;

        const fetchBriefing = async () => {
            try {
                setIsFetchingBriefing(true);
                const { aiGenerateBriefing } = await import("@/api");
                const res = await aiGenerateBriefing(bodyText || "", [], {}, ctx.conversationId);
                if (res.ok) {
                    setBriefing(res.summary);
                }
            } catch (err) {
                console.error("Erro ao obter briefing:", err);
            } finally {
                setIsFetchingBriefing(false);
            }
        };

        fetchBriefing();
    }, [ctx.conversationId, ctx.isCompose, bodyText]);

    // Extract contacts from body text when it changes
    useEffect(() => {
        if (!bodyText || ctx.isCompose) return;

        const extractContacts = async () => {
            try {
                const res = await aiGenerate({
                    action: "extract_contacts" as any,
                    mode: "fast",
                    locale: "pt-PT",
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
                    const emails = res.text.split(";").map(e => e.trim()).filter(Boolean);
                    setSuggestedContacts(emails);
                }
            } catch (err) {
                console.error("Erro ao extrair contactos:", err);
            }
        };
        extractContacts();
    }, [bodyText, ctx.isCompose]);

    // Sync draft defaults from context
    useEffect(() => {
        setDraftTo((ctx.toRecipients || []).map((r: any) => r.email));
        setDraftCc((ctx.ccRecipients || []).map((r: any) => r.email));
        setDraftSubject(ctx.subject || "");
    }, [ctx]);

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
                setPrompt((prev: string) => {
                    const updated = (prev + " " + newFinal).trim();
                    setAiState({ prompt: updated });
                    return updated;
                });
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
                inputText: extraPrompt || prompt,
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
            });

            if (res.ok) {
                setAiState({ action, output: res.text });
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
                    { role: "user" as const, content: extraPrompt || prompt },
                    { role: "assistant" as const, content: fullText }
                ].slice(-4);
                setAiState({ output: fullText, history: newHistory });
                setPrompt("");
                setShowDraftPreview(true);
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

            if (ctx.isCompose) {
                setDebugLog("A atualizar rascunho...");

                // Sync metadata first
                await setRecipients("to", draftTo);
                await setRecipients("cc", draftCc);
                await setSubject(draftSubject);

                // Insert body
                await insertTextToBody(output);

                setDebugLog("Atualizado com sucesso!");
                setMsg("Draft atualizado com sucesso!");
                setTimeout(() => setMsg(""), 3000);
                return;
            }

            // If not in compose mode, try to open a Draft based on action
            setDebugLog("A abrir rascunho (não é modo edição)...");
            if (aiState.action === "reply") {
                await displayReplyForm(output);
            } else if (aiState.action === "forward") {
                await displayForwardForm(output);
            } else {
                throw new Error("Para inserir este texto, tens de estar em modo de edição ou abrir uma resposta.");
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

    const [activeMenu, setActiveMenu] = useState<"lang" | "mode" | "presets" | null>(null);
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
        setAiState({ history: [], output: "" });
        setPrompt("");
        clearFiles(); // Using the new context helper
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


    return (
        <div style={S.container}>
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
            {briefing && !isFetchingBriefing && (
                <div style={S.briefingCard}>
                    <div style={S.briefingHeader}>
                        <div style={{ display: "flex", alignItems: "center", gap: "6px" }}>
                            <Icons.Sparkles size={12} color="#1e40af" />
                            <span>30-Second Briefing</span>
                        </div>
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

            {isFetchingBriefing && (
                <div style={{ ...S.briefingCard, background: "rgba(0,0,0,0.02)", borderStyle: "dashed" }}>
                    <div style={S.skeletonText}>
                        <Icons.RotateCcw size={10} style={{ animation: "spin 1s linear infinite" }} />
                        A gerar briefing do thread...
                    </div>
                </div>
            )}

            {debugLog && (
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
                        aiState.smartReplies.map((intent: string, idx: number) => (
                            <button
                                key={idx}
                                onClick={() => {
                                    setPrompt(intent);
                                    handleGenerate("reply", intent);
                                }}
                                style={S.intentChip}
                            >
                                {intent}
                            </button>
                        ))
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

            {/* AI Task Extraction Review */}
            {showTaskReview && extractedTasks.length > 0 && (
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
                                    Ignorar
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
                        style={{ ...S.secondaryBtnLink, width: "68px", minWidth: "68px", justifyContent: "flex-start", padding: "0 6px" }}
                        onClick={() => setActiveMenu(activeMenu === "lang" ? null : "lang")}
                        title="Idioma"
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
                        <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 800 }}>
                            {(aiState.locale || "auto").split("-")[0].toUpperCase()}
                        </span>
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
                                    <span style={{ fontSize: "9px", fontWeight: 800 }}>
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
                        style={{ ...S.secondaryBtnLink, width: "84px", minWidth: "84px", justifyContent: "flex-start", padding: "0 8px" }}
                        onClick={() => setActiveMenu(activeMenu === "presets" ? null : "presets")}
                        title="Modelos de Resposta Rápidos"
                    >
                        <Icons.Settings size={11} style={{ opacity: 0.6 }} />
                        <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 800 }}>MODELOS</span>
                    </button>

                    {activeMenu === "presets" && (
                        <div style={{ ...S.cascadeMenu, width: "160px" }}>
                            {settings?.responsePresets?.length > 5 && (
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
                            )}

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
                                        <span style={{ fontWeight: 800, fontSize: "10px" }}>{p.name.toUpperCase()}</span>
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

                {/* Mode Cascade */}
                <div style={{ position: "relative" }}>
                    <button
                        className="iccc-glossy-pill iccc-secondary-pill"
                        style={{ ...S.secondaryBtnLink, width: "68px", minWidth: "68px", justifyContent: "flex-start", padding: "0 6px" }}
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
                        <span style={{ fontSize: "9px", marginLeft: "4px", fontWeight: 800 }}>MODO</span>
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
                                    <span style={{ fontWeight: 800 }}>{r.label.toUpperCase()}</span>
                                </button>


                            ))}
                        </div>
                    )}
                </div>
            </div>

            {
                (output || isGenerating || aiState.history.length > 0) && (
                    <div style={S.outputCard}>
                        <div style={S.outputHeader}>
                            <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                                <span>Sugestão da IA</span>
                                {isGenerating && <div style={S.typingDots}><span>.</span><span>.</span><span>.</span></div>}
                            </div>
                            <div style={{ display: "flex", gap: "8px" }}>
                                <button style={S.actionBtn} onClick={handleResetConversation} title="Limpar conversa (Reset Total)">
                                    <Icons.Trash size={14} />
                                </button>
                                <button style={S.actionBtn} onClick={handleExport} title="Descarregar Texto (.txt)">
                                    <Icons.Download size={14} />
                                </button>
                                <button style={S.actionBtn} onClick={() => navigator.clipboard.writeText(output)} title="Copiar texto">
                                    <Icons.Clipboard size={14} />
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
                                    title="Agendar Reunião"
                                >
                                    <Icons.Calendar size={14} style={{ marginRight: "4px" }} />
                                    Agendar
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
                                    title="Inserir no Email"
                                >
                                    <Icons.Send size={14} style={{ marginRight: "4px" }} />
                                    {ctx.isCompose ? "Atualizar" : "Inserir"}
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

                                    {settings?.contactAliases?.length > 0 && (
                                        <div style={{ marginTop: "8px" }}>
                                            <span style={{ fontSize: "9px", fontWeight: 800, color: "#1e40af", textTransform: "uppercase" }}>Atalhos Rápidos (Contactos):</span>
                                            <div style={{ display: "flex", flexWrap: "wrap", gap: "4px", marginTop: "4px" }}>
                                                {settings.contactAliases.map((c: any) => (
                                                    <button
                                                        key={c.id}
                                                        style={S.suggestedChip}
                                                        onClick={() => {
                                                            if (!draftTo.includes(c.email)) setDraftTo([...draftTo, c.email]);
                                                        }}
                                                        title={c.email}
                                                    >
                                                        <Icons.User size={10} style={{ marginRight: "3px" }} />
                                                        {c.name}
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    )}

                                    {suggestedContacts.length > 0 && (
                                        <div style={{ marginTop: "8px" }}>
                                            <span style={{ fontSize: "9px", fontWeight: 800, color: "#1e40af", textTransform: "uppercase" }}>Contactos Detetados no Email:</span>
                                            <div style={{ display: "flex", flexWrap: "wrap", gap: "4px", marginTop: "4px" }}>
                                                {suggestedContacts.map(email => (
                                                    <button
                                                        key={email}
                                                        style={S.suggestedChip}
                                                        onClick={() => {
                                                            if (!draftTo.includes(email)) setDraftTo([...draftTo, email]);
                                                        }}
                                                    >
                                                        + {email}
                                                    </button>
                                                ))}
                                            </div>
                                        </div>
                                    )}
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
                                value={prompt}
                                onChange={(e) => setPrompt(e.target.value)}
                                onKeyDown={(e) => handleKeyDown(e, "refine")}
                                rows={1}
                            />
                            <button
                                disabled={isGenerating || !prompt}
                                onClick={() => handleGenerate("refine")}
                                style={{
                                    ...S.chatSendBtn,
                                    opacity: !prompt || isGenerating ? 0.5 : 1
                                }}
                            >
                                <Icons.Sparkles size={14} />
                            </button>
                        </div>
                    </div>
                )
            }
        </div >
    );
};

const S: Record<string, React.CSSProperties> = {
    primaryBtnPill: {
        boxSizing: "border-box",
        height: "26px", minHeight: "26px", maxHeight: "26px",
        borderRadius: "16px",
        border: "1px solid rgba(0, 80, 180, 0.4)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "4px",
        padding: "0 10px",
        fontSize: "10px",
        fontWeight: 800,
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
        width: "28px", minWidth: "28px", maxWidth: "28px",
        height: "26px", minHeight: "26px", maxHeight: "26px",
        borderRadius: "16px",
        border: "1px solid rgba(200, 210, 230, 0.6)",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        gap: "5px",
        padding: "0",
        fontSize: "10px",
        fontWeight: 800,
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
        height: "26px", minHeight: "26px", maxHeight: "26px",
        borderRadius: "16px",
        border: "1px solid rgba(200, 210, 230, 0.6)",
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
        zIndex: 500,
        background: "transparent",
        width: "fit-content"
    },
    cascadeItem: {
        boxSizing: "border-box",
        height: "26px", minHeight: "26px",
        minWidth: "80px",
        whiteSpace: "nowrap",
        borderRadius: "16px",
        border: "1px solid rgba(200, 210, 230, 0.6)",
        backdropFilter: "blur(12px)",
        WebkitBackdropFilter: "blur(12px)",
        display: "flex",
        alignItems: "center",
        justifyContent: "flex-start",
        gap: "6px",
        padding: "0 10px",
        fontSize: "9px",
        fontWeight: 800,
        lineHeight: 1,
        textTransform: "uppercase",
        cursor: "pointer",
        background: "linear-gradient(180deg, rgba(255,255,255,0.98) 0%, rgba(240,245,255,0.95) 100%)",
        color: "#172B4D",
        boxShadow: "0 4px 12px rgba(0,0,0,0.12), inset 0 1px 0 rgba(255,255,255,1)",
        transition: "all 0.18s ease",
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
        gap: "4px",
        paddingBottom: "4px",
        alignItems: "center",
        overflow: "visible"
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
        fontSize: "10px",
        fontWeight: 700,
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
        width: "28px",
        height: "22px",
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
        fontWeight: 800,
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
        fontWeight: 800,
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
