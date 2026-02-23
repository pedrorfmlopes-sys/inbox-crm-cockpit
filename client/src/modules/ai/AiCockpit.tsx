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

    const localeOptions: Array<{ label: string; value: AiLocale; icon: string }> = [
        { label: "Auto", value: "auto", icon: "🤖" },
        { label: "Português", value: "pt-PT", icon: "🇵🇹" },
        { label: "English", value: "en-GB", icon: "🇬🇧" },
        { label: "Español", value: "es-ES", icon: "🇪🇸" },
        { label: "Italiano", value: "it-IT", icon: "🇮🇹" },
        { label: "Deutsch", value: "de-DE", icon: "🇩🇪" },
    ];

    return (
        <div style={S.container}>
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
                    <div style={{ display: "flex", gap: "8px" }}>
                        <button
                            style={{
                                ...S.secondaryBtn,
                                borderColor: isRecording ? "#ef4444" : "var(--iccc-card-border)",
                                color: isRecording ? "#ef4444" : "var(--iccc-text-muted)",
                                background: isRecording ? "rgba(239, 68, 68, 0.1)" : "var(--iccc-bg)",
                            }}
                            onClick={toggleRecording}
                            title="Ditado por voz"
                        >
                            <Icons.Microphone size={14} />
                        </button>
                        <div style={{ width: "1px", height: "24px", background: "var(--iccc-card-border)", margin: "0 4px" }}></div>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("summarize")}
                            disabled={isGenerating}
                            title={files.length > 0 ? "Resumir email e anexos identificados" : "Resumir este email"}
                        >
                            <Icons.Receipt size={14} />
                        </button>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("tasks")}
                            disabled={isGenerating}
                            title="Extrair tarefas"
                        >
                            <Icons.Check size={14} />
                        </button>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("forward")}
                            disabled={isGenerating}
                            title="Reenviar (Rascunho)"
                        >
                            <Icons.Send size={14} />
                        </button>
                        <button
                            style={{
                                ...S.secondaryBtn,
                                color: isImporting ? "var(--iccc-pill-active-bg)" : "var(--iccc-text-muted)"
                            }}
                            onClick={handleImportAttachments}
                            disabled={isImporting}
                            title="Importar anexos deste email"
                        >
                            {isImporting ? <Icons.RotateCcw size={14} style={{ animation: "spin 1s linear infinite" }} /> : <Icons.Paperclip size={14} />}
                        </button>
                    </div>
                    <button
                        style={S.generateBtn}
                        onClick={() => handleGenerate("reply")}
                        disabled={isGenerating}
                    >
                        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                            {isGenerating ? "A gerar..." : "Gerar Resposta"}
                            <Icons.Sparkles size={14} />
                        </div>
                    </button>
                </div>
            </div>

            <div style={S.refinerRow}>
                <div style={{ display: "flex", alignItems: "center", background: "var(--iccc-card-bg)", border: "1px solid var(--iccc-card-border)", borderRadius: "4px", padding: "0 2px", marginRight: "3px", height: "24px" }}>
                    <select
                        style={S.langSelect}
                        title="Selecionar idioma de resposta"
                        value={aiState.locale || "auto"}
                        onChange={(e) => {
                            const val = e.target.value as AiLocale;
                            setAiState({ locale: val });
                            if (output) handleGenerate("rewrite", output);
                        }}
                    >
                        {localeOptions.map((opt) => (
                            <option key={opt.value} value={opt.value}>
                                {opt.icon} {opt.label}
                            </option>
                        ))}
                    </select>
                </div>

                <div style={{ width: "1px", height: "20px", background: "var(--iccc-card-border)", margin: "0 4px" }}></div>

                {toneRefiners.map((r) => (
                    <button
                        key={r.label}
                        style={{
                            ...S.refinerChip,
                            borderColor: aiState.tone === r.tone ? "var(--iccc-pill-active-bg)" : "var(--iccc-card-border)",
                            background: aiState.tone === r.tone ? "var(--iccc-pill-active-bg)" : "transparent",
                            color: aiState.tone === r.tone ? "white" : "var(--iccc-text)",
                        }}
                        onClick={() => {
                            setAiState({ tone: r.tone });
                            if (output) handleGenerate("rewrite", output);
                        }}
                    >
                        <span style={{ marginRight: "4px" }}>{r.icon}</span>
                        {r.label}
                    </button>
                ))}
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
    // ... container etc
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
    generateBtn: {
        background: "linear-gradient(90deg, #3b82f6 0%, #2563eb 100%)",
        color: "white",
        border: "none",
        borderRadius: "6px",
        padding: "5px 12px",
        fontSize: "11px",
        fontWeight: 600,
        cursor: "pointer",
        boxShadow: "0 2px 4px rgba(37, 99, 235, 0.15)",
        transition: "transform 0.1s ease",
    },
    secondaryBtn: {
        background: "transparent",
        color: "var(--iccc-text-muted)",
        border: "1px solid transparent",
        borderRadius: "8px",
        padding: "4px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        transition: "all 0.2s",
    },
    refinerRow: {
        display: "flex",
        gap: "4px",
        overflowX: "auto",
        paddingBottom: "4px",
        alignItems: "center",
    },
    refinerChip: {
        flexShrink: 0,
        padding: "2px 4px",
        borderRadius: "4px",
        border: "1px solid var(--iccc-card-border)",
        background: "rgba(0,0,0,0.02)",
        fontSize: "10px",
        fontWeight: 600,
        cursor: "pointer",
        transition: "all 0.2s",
        display: "flex",
        alignItems: "center",
        height: "24px",
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
    outputActions: {
        display: "flex",
        justifyContent: "flex-end",
        gap: "16px",
        borderTop: "1px solid rgba(0,0,0,0.05)",
        paddingTop: "16px",
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
    langSelect: {
        background: "transparent",
        border: "none",
        color: "var(--iccc-text)",
        fontSize: "10px",
        height: "22px",
        padding: "0",
        outline: "none",
        cursor: "pointer",
        fontWeight: 700,
        maxWidth: "60px",
    },
};
