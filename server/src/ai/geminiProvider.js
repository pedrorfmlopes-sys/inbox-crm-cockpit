// server/src/ai/geminiProvider.js
import fs from "fs";
import path from "path";
// Uses native fetch to call Google Generative AI API (Gemini).
// Supports multimodal inputs (text + files).

export async function geminiCreateResponse({
    apiKey,
    model = "gemini-1.5-flash",
    instructions,
    input,
    files = [],
    history = [], // NEW: Support for conversation history
    max_output_tokens = 2048,
    temperature = 0.2,
}) {
    if (!apiKey) throw Object.assign(new Error("GEMINI_API_KEY em falta"), { status: 400 });

    const url = `https://generativelanguage.googleapis.com/v1beta/models/${model}:generateContent?key=${apiKey}`;
    const logPath = "gemini-debug.log";
    fs.appendFileSync(logPath, `\n[${new Date().toISOString()}] Calling: ${url.replace(apiKey, 'REDACTED')}\n`);

    // Build the prompt parts
    const parts = [];

    // 1. System instructions (as text)
    if (instructions) {
        parts.push({ text: `SYSTEM INSTRUCTIONS:\n${instructions}\n\n` });
    }

    // 2. Attached files
    for (const f of files) {
        if (f.content && f.type) {
            const base64Data = f.content.includes(",") ? f.content.split(",")[1] : f.content;
            fs.appendFileSync(logPath, `[gemini] File: ${f.name}, Type: ${f.type}, B64 Start: ${base64Data.slice(0, 30)}..., Total: ${base64Data.length}\n`);
            parts.push({
                inlineData: {
                    mimeType: f.type,
                    data: base64Data,
                }
            });
            parts.push({ text: `DOC_CONTENT_${f.name}:\n` }); // Hint for the AI
        }
    }

    // 3. User input
    parts.push({ text: input || "ok" });

    const contents = [];

    // Add history (Gemini format: user | model)
    for (const h of history) {
        contents.push({
            role: h.role === "user" ? "user" : "model",
            parts: [{ text: h.content }]
        });
    }

    // Add current user turn
    contents.push({ parts });

    const body = {
        contents,
        generationConfig: {
            maxOutputTokens: max_output_tokens,
            temperature: temperature,
        },
    };

    try {
        const res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });

        const data = await res.json();

        if (!res.ok) {
            let msg = data?.error?.message || `Gemini HTTP ${res.status}`;

            // Translate common errors for better UX
            if (res.status === 429) msg = "Quota Excedida (Gemini). Tenta novamente em breve.";
            if (res.status === 401 || res.status === 403) msg = "Chave de API Gemini Inválida ou Expirada.";
            if (res.status === 404) msg = "Modelo Gemini não encontrado ou indisponível.";

            const err = new Error(msg);
            err.status = res.status;
            err.details = data;
            throw err;
        }

        const outputText = data?.candidates?.[0]?.content?.parts?.[0]?.text || "";
        return { raw: data, text: outputText.trim() };
    } catch (e) {
        console.error("[ai] Gemini provider error:", e.message);
        throw e;
    }
}

export async function geminiListModels(apiKey) {
    if (!apiKey) throw new Error("GEMINI_API_KEY em falta");
    const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${apiKey}`;
    const res = await fetch(url);
    if (!res.ok) throw new Error(`Gemini Models HTTP ${res.status}`);
    const data = await res.json();
    // Filter for generative models that support generateContent
    return (data.models || [])
        .filter(m => m.supportedGenerationMethods.includes("generateContent"))
        .map(m => m.name.replace("models/", ""))
        .sort();
}
