import React, { useState, useEffect, useRef } from "react";
import { useCockpit } from "@/components/shell/CockpitProvider";
import { aiGenerate, type AiAction, type AiTone, type AiLocale } from "@/ai/aiClient";
import { insertTextToBody, isComposeMode, displayReplyForm, displayForwardForm } from "@/office";
import { getSettings } from "@/settings";
import * as Icons from "@/ui/icons";

export const AiCockpit: React.FC = () => {
    const { ctx, bodyText, setMsg, aiState, setAiState, files } = useCockpit();

    // Local state for immediate typing feel
    // Initialized from context, but NOT synced on every render to avoid loops
    const [prompt, setPrompt] = useState(aiState.prompt);
    const [output, setOutput] = useState(aiState.output);
    const [isGenerating, setIsGenerating] = useState(false);
    const [debugLog, setDebugLog] = useState("");

    // Voice Dictation State
    const [isRecording, setIsRecording] = useState(false);
    const recognitionRef = useRef<any>(null);

    // CRITICAL: Only sync local state from context when the conversation (email) changes.
    // This prevents the "typing -> context update -> local reset" loop.
    useEffect(() => {
        setPrompt(aiState.prompt);
        setOutput(aiState.output);
        setDebugLog(""); // Clear debug log on switch
    }, [ctx.conversationId]); // DEPENDENCY on conversationId ensures reset on switch

    const handlePromptChange = (val: string) => {
        setPrompt(val);
        // Sync to context so it persists if we switch tabs/emails and come back
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
        recognition.lang = aiState.locale === "auto" ? "pt-PT" : aiState.locale;
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
            let interimTranscript = "";
            let newFinal = "";

            for (let i = event.resultIndex; i < event.results.length; ++i) {
                if (event.results[i].isFinal) {
                    newFinal += event.results[i][0].transcript;
                } else {
                    interimTranscript += event.results[i][0].transcript;
                }
            }

            // Only update if we have new FINAL content to append
            // We ignore interim results for the prompt update to avoid flickering/loops
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

    async function handleGenerate(action: AiAction = "rewrite", extraPrompt?: string) {
        if (!ctx.subject && !prompt && !extraPrompt) {
            setMsg("Escreve algo ou seleciona um email primeiro.");
            return;
        }

        setIsGenerating(true);
        setOutput("");
        setDebugLog("");

        try {
            const settings = await getSettings();
            // 1. Prepare files to send (Raw Base64)
            const filesToSend = files || [];

            const isRefining = action === "rewrite" || action === "refine";
            const historyToSend = isRefining ? aiState.history : [];

            // If we are starting a NEW task (not refining), clear previous history
            if (!isRefining && aiState.history.length > 0) {
                setAiState({ history: [] });
            }

            const res = await aiGenerate({
                action,
                mode: "fast", // TODO: make dynamic?
                tone: settings.tone || "neutro",
                locale: settings.replyLanguage || "pt-PT",
                inputText: extraPrompt || prompt,
                files: filesToSend, // Send raw files for native AI processing
                history: historyToSend,
                email: {
                    subject: ctx.subject || "",
                    from: ctx.from || "",
                    to: ctx.to || [],
                    cc: ctx.cc || [],
                    bodyText: bodyText || "",
                },
                knowledge: settings.aiKnowledge || [],
            });

            if (res.ok) {
                setAiState({ action, output: res.text }); // Store action for smart insert
                // Efeito de streaming simulado para estética
                let fullText = res.text;
                let current = "";
                const words = fullText.split(" ");
                for (let i = 0; i < words.length; i++) {
                    current += words[i] + " ";
                    setOutput(current);
                    // if (i % 5 === 0) setAiState({ output: current }); // batch context updates
                    await new Promise((r) => setTimeout(r, 20));
                }

                // Add to history (Cap history to avoid bloat)
                const newHistory = [
                    ...(isRefining ? aiState.history : []),
                    { role: "user" as const, content: extraPrompt || prompt },
                    { role: "assistant" as const, content: fullText }
                ].slice(-4); // Keep last 2 turns
                setAiState({ output: fullText, history: newHistory });
                setPrompt(""); // Clear prompt after generation for follow-up
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

    const refiners: Array<{ label: string; tone?: AiTone; locale?: AiLocale; icon: React.ReactNode }> = [
        { label: "Profissional", tone: "formal", icon: <Icons.Building size={14} /> },
        { label: "Amigável", tone: "simpático", icon: <Icons.Sparkles size={14} /> },
        { label: "Curto", tone: "curto", icon: <Icons.Receipt size={14} /> },
        { label: "English", locale: "en-GB", icon: <span style={{ fontSize: "12px" }}>🇬🇧</span> },
        { label: "Português", locale: "pt-PT", icon: <span style={{ fontSize: "12px" }}>🇵🇹</span> },
    ];

    return (
        <div style={S.container}>
            {debugLog && (
                <div style={{ padding: "8px", background: "#fee2e2", color: "#b91c1c", fontSize: "11px", borderRadius: "4px", border: "1px solid #fca5a5" }}>
                    DEBUG: {debugLog}
                </div>
            )}
            <div style={S.inputCard}>
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
                            <Icons.Microphone size={16} />
                        </button>
                        <div style={{ width: "1px", height: "24px", background: "var(--iccc-card-border)", margin: "0 4px" }}></div>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("summarize")}
                            disabled={isGenerating}
                            title={files.length > 0 ? "Resumir email e anexos identificados" : "Resumir este email"}
                        >
                            <Icons.Receipt size={16} />
                        </button>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("tasks")}
                            disabled={isGenerating}
                            title="Extrair tarefas"
                        >
                            <Icons.Check size={16} />
                        </button>
                        <button
                            style={S.secondaryBtn}
                            onClick={() => handleGenerate("forward")}
                            disabled={isGenerating}
                            title="Reenviar (Rascunho)"
                        >
                            <Icons.Send size={16} />
                        </button>
                    </div>
                    <button
                        style={S.generateBtn}
                        onClick={() => handleGenerate("reply")}
                        disabled={isGenerating}
                    >
                        <div style={{ display: "flex", alignItems: "center", gap: "8px" }}>
                            {isGenerating ? "A gerar..." : "Gerar Resposta"}
                            <Icons.Sparkles size={16} />
                        </div>
                    </button>
                </div>
            </div>

            <div style={S.refinerRow}>
                {refiners.map((r) => (
                    <button
                        key={r.label}
                        style={{
                            ...S.refinerChip,
                            borderColor: aiState.tone === r.tone || aiState.locale === r.locale ? "var(--iccc-pill-active-bg)" : "var(--iccc-card-border)",
                            background: aiState.tone === r.tone || aiState.locale === r.locale ? "var(--iccc-pill-active-bg)" : "transparent",
                            color: aiState.tone === r.tone || aiState.locale === r.locale ? "white" : "var(--iccc-text)",
                        }}
                        onClick={() => {
                            if (r.tone) setAiState({ tone: r.tone });
                            if (r.locale) setAiState({ locale: r.locale });
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
        height: "28px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
    },
    container: {
        display: "flex",
        flexDirection: "column",
        gap: "8px",
        paddingTop: "2px",
    },
    inputCard: {
        background: "var(--iccc-card-bg)",
        border: "1px solid var(--iccc-card-border)",
        borderRadius: "12px",
        padding: "10px 12px",
        boxShadow: "0 1px 4px rgba(0,0,0,0.03)",
        display: "flex",
        flexDirection: "column",
        gap: "6px",
    },
    textarea: {
        width: "100%",
        minHeight: "80px",
        background: "transparent",
        border: "none",
        color: "var(--iccc-text)",
        fontFamily: "var(--iccc-font)",
        fontSize: "13px",
        resize: "none",
        outline: "none",
    },
    inputFooter: {
        display: "flex",
        justifyContent: "space-between",
        alignItems: "center",
    },
    generateBtn: {
        background: "var(--iccc-btn-bg)",
        color: "var(--iccc-btn-text)",
        border: "none",
        borderRadius: "8px",
        padding: "8px 16px",
        fontSize: "12px",
        fontWeight: 600,
        cursor: "pointer",
    },
    secondaryBtn: {
        background: "transparent",
        color: "var(--iccc-text-muted)",
        border: "1px solid transparent",
        borderRadius: "8px",
        padding: "6px",
        display: "flex",
        alignItems: "center",
        justifyContent: "center",
        cursor: "pointer",
        transition: "all 0.2s",
    },
    refinerRow: {
        display: "flex",
        gap: "8px",
        overflowX: "auto",
        paddingBottom: "4px",
    },
    refinerChip: {
        flexShrink: 0,
        padding: "4px 10px",
        borderRadius: "6px",
        border: "1px solid var(--iccc-card-border)",
        background: "rgba(0,0,0,0.03)",
        fontSize: "11px",
        fontWeight: 500,
        cursor: "pointer",
        transition: "all 0.2s",
        display: "flex",
        alignItems: "center",
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
        lineHeight: "1.5",
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
};
