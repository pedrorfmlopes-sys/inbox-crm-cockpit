import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { aiGenerate, type AiAction, type AiTone, type AiLocale } from "@/ai/aiClient";
import { insertTextToBody, isComposeMode, displayReplyForm, displayForwardForm, displayNewMeetingForm } from "@/office";
import { getSettings } from "@/settings";
import * as Icons from "@/ui/icons";

export const AiCockpit: React.FC = () => {
    const { ctx, bodyText, setMsg, aiState, setAiState, files, addFile } = useCockpit();

    // Local state for immediate typing feel
    // Initialized from context, but NOT synced on every render to avoid loops
    const [prompt, setPrompt] = useState(aiState.prompt);
    const [output, setOutput] = useState(aiState.output);
    const [isGenerating, setIsGenerating] = useState(false);
    const [isFetchingIntents, setIsFetchingIntents] = useState(false);
    const [isImporting, setIsImporting] = useState(false);
    const [debugLog, setDebugLog] = useState("");

    // Voice Dictation State
    const [isRecording, setIsRecording] = useState(false);
    const recognitionRef = useRef<any>(null);

    // CRITICAL: Only sync local state from context when the conversation (email) changes.
    useEffect(() => {
        setPrompt(aiState.prompt);
        setOutput(aiState.output);
        setDebugLog(""); // Clear debug log on switch
    }, [ctx.conversationId]);

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
                    action: "intent_proposals",
                    mode: "fast",
                    locale: (settings.replyLanguage || "pt-PT") as any,
                    tone: settings.tone || "neutro",
                    email: {
                        subject: ctx.subject || "",
                        from: ctx.fromEmail || "",
                        to: (ctx.toRecipients || []).map((r: any) => r.email),
                        cc: (ctx.ccRecipients || []).map((r: any) => r.email),
                        bodyText: bodyText || ""
                    },
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
                email: {
                    subject: ctx.subject || "",
                    from: ctx.fromEmail || "",
                    to: (ctx.toRecipients || []).map(r => r.email),
                    cc: (ctx.ccRecipients || []).map(r => r.email),
                    bodyText: bodyText || ""
                },
                persona: {
                    userRole: settings.userRole,
                    styleContext: settings.styleContext,
                    styleExamples: settings.styleExamples,
                },
                history: isRefining ? aiState.history : [],
                knowledge: settings.aiKnowledge || [],
            });

            if (res.ok) {
                setAiState({ action, output: res.text });
                let fullText = res.text;
                let current = "";
                const words = fullText.split(" ");
                for (let i = 0; i < words.length; i++) {
                    current += words[i] + " ";
                    setOutput(current);
                    await new Promise((r) => setTimeout(r, 20));
                }

                const newHistory = [
                    ...(isRefining ? aiState.history : []),
                    { role: "user" as const, content: extraPrompt || prompt },
                    { role: "assistant" as const, content: fullText }
                ].slice(-4);
                setAiState({ output: fullText, history: newHistory });
                setPrompt("");
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
                setDebugLog("A inserir texto...");
                await insertTextToBody(output);
                setDebugLog("Inserido com sucesso!");
                setMsg("Texto inserido com sucesso!");
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

    const [activeMenu, setActiveMenu] = useState<"lang" | "mode" | null>(null);
    const menuRef = useRef<HTMLDivElement>(null);

    useEffect(() => {
        const handleClickOutside = (e: MouseEvent) => {
            if (menuRef.current && !menuRef.current.contains(e.target as Node)) setActiveMenu(null);
        };
        document.addEventListener("mousedown", handleClickOutside);
        return () => document.removeEventListener("mousedown", handleClickOutside);
    }, []);

    const handleKeyDown = (e: React.KeyboardEvent, action: AiAction = "reply") => {
        if (e.key === "Enter" && (e.ctrlKey || e.metaKey)) {
            e.preventDefault();
            handleGenerate(action);
        }
        if (e.key === "Escape") {
            setPrompt("");
        }
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
                        >

                            <Icons.Receipt size={12} />
                        </button>
                        <button
                            className="iccc-glossy-pill iccc-secondary-pill"
                            style={S.secondaryBtnPill}
                            onClick={() => handleGenerate("tasks")}
                            disabled={isGenerating}
                            title="Extrair tarefas"
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
                        <div style={S.cascadeMenu}>
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
                        <div style={S.cascadeMenu}>
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

            {(output || isGenerating || aiState.history.length > 0) && (
                <div style={S.outputCard}>
                    <div style={S.outputHeader}>
                        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                            <span>Sugestão da IA</span>
                            {isGenerating && <div style={S.typingDots}><span>.</span><span>.</span><span>.</span></div>}
                        </div>
                        <div style={{ display: "flex", gap: "8px" }}>
                            {aiState.history.length > 0 && (
                                <button style={S.actionBtn} onClick={() => setAiState({ history: [] })} title="Limpar conversa">
                                    <Icons.Trash size={14} />
                                </button>
                            )}
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
                                <Icons.ExternalLink size={14} style={{ marginRight: "4px" }} />
                                Inserir
                            </button>
                        </div>
                    </div>

                    {/* Iterative Result Area */}
                    <div style={{ position: "relative" }}>
                        <div style={S.outputText}>
                            {output}
                        </div>
                        {isGenerating && !output && (
                            <div style={{ padding: "20px 0", color: "var(--iccc-text-muted)", fontStyle: "italic" }}>
                                A pensar...
                            </div>
                        )}
                    </div>

                    {/* Quick Refinement Input */}
                    <div style={S.chatInputWrapper}>
                        <input
                            style={S.chatInput}
                            placeholder="Refinar resposta (ex: faz mais curto)..."
                            value={prompt}
                            onChange={(e) => setPrompt(e.target.value)}
                            onKeyDown={(e) => handleKeyDown(e, "refine")}
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
            )}
        </div>
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
        left: 0,
        display: "flex",
        flexDirection: "column",
        gap: "4px",
        zIndex: 100,
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
};
