// server/src/ai/geminiProvider.js
import fs from "fs";
import path from "path";
// Uses native fetch to call Google Generative AI API (Gemini).
// Supports multimodal inputs (text + files).

const GEMINI_ALLOWLIST = ["gemini-1.5-flash", "gemini-1.5-pro", "gemini-2.0-flash", "gemini-2.0-flash-exp", "gemini-2.0-pro-exp"];

function sanitizeGeminiModel(m) {
    if (!m) return "gemini-1.5-flash";
    // Remove anything in parentheses (e.g. "gemini-2.0-flash (NextGen)" -> "gemini-2.0-flash")
    let clean = String(m).split("(")[0].trim();
    // Also take only first word if spaces exist outside parentheses
    clean = clean.split(" ")[0].trim();

    if (GEMINI_ALLOWLIST.includes(clean)) return clean;
    // Basic prefix check if it starts with gemini- but not in allowlist
    if (clean.startsWith("gemini-")) return clean;

    return "gemini-1.5-flash";
}

export async function geminiCreateResponse({
    apiKey,
    model: requestedModel = "gemini-1.5-flash",
    instructions,
    input,
    files = [],
    history = [], // NEW: Support for conversation history
    max_output_tokens = 2048,
    temperature = 0.2,
}) {
    if (!apiKey) throw Object.assign(new Error("GEMINI_API_KEY em falta"), { status: 400 });

    const sanitizedModel = sanitizeGeminiModel(requestedModel);
    let effectiveModel = sanitizedModel;

    const buildUrl = (m) => `https://generativelanguage.googleapis.com/v1beta/models/${m}:generateContent?key=${apiKey}`;
    let url = buildUrl(effectiveModel);

    const logPath = "gemini-debug.log";
    fs.appendFileSync(logPath, `\n[${new Date().toISOString()}] Calling: ${url.replace(apiKey, 'REDACTED')} (Requested: ${requestedModel})\n`);

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
        let res = await fetch(url, {
            method: "POST",
            headers: { "Content-Type": "application/json" },
            body: JSON.stringify(body),
        });

        let data = await res.json();

        // Specific Fallback for 404 (Model not found)
        if (res.status === 404 && effectiveModel !== "gemini-1.5-flash") {
            const oldModel = effectiveModel;
            effectiveModel = "gemini-1.5-flash";
            url = buildUrl(effectiveModel);
            console.log(`[ai] Gemini 404 for ${oldModel}. Falling back to ${effectiveModel}...`);
            fs.appendFileSync(logPath, `[${new Date().toISOString()}] 404 Fallback triggered. New URL: ${url.replace(apiKey, 'REDACTED')}\n`);

            res = await fetch(url, {
                method: "POST",
                headers: { "Content-Type": "application/json" },
                body: JSON.stringify(body),
            });
            data = await res.json();
        }

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
        return {
            raw: data,
            text: outputText.trim(),
            requestedModel,
            sanitizedModel,
            effectiveModel,
            providerUsed: "gemini"
        };
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
