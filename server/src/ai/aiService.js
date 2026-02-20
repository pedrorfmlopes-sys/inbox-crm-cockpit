import { getAiConfig } from "./aiConfig.js";
import { openaiCreateResponse } from "./openaiProvider.js";
import { geminiCreateResponse } from "./geminiProvider.js";
import { extractTextFromPdfBuffer } from "./pdfHelper.js";

export function getAiMeta() {
  const cfg = getAiConfig();
  return {
    enabled: cfg.enabled,
    provider: cfg.provider,
    keyPresent:
      cfg.provider === "openai"
        ? Boolean(cfg.openai.apiKey)
        : cfg.provider === "gemini"
          ? Boolean(cfg.gemini.apiKey)
          : false,
    openaiModelFast: cfg.openai.modelFast,
    openaiModelQuality: cfg.openai.modelQuality,
    geminiModel: cfg.gemini.model,
  };
}

export async function aiCreateText({
  mode = "fast",
  instructions,
  input,
  files = [], // NEW: Support for raw files (multimodal)
  history = [], // NEW: Support for conversation history
  max_output_tokens = 256,
  temperature = 0.2,
  isFallback = false, // Internal flag to prevent infinite loops
}) {
  const cfg = getAiConfig();

  if (!cfg.enabled) {
    throw Object.assign(new Error("AI desativado."), { status: 400 });
  }

  // --- HYBRID PROVIDER SELECTION ---
  // If we have files (PDFs/Images) and a Gemini key, we use Gemini (it's best for this).
  // Otherwise, we use the default provider (usually OpenAI).
  let selectedProvider = cfg.provider;

  if (files.length > 0 && cfg.gemini.apiKey) {
    selectedProvider = "gemini";
    console.log(`[ai] Switch to Gemini for multimodal request (found ${files.length} files)`);
  }

  try {
    if (selectedProvider === "openai") {
      if (!cfg.openai.apiKey) throw Object.assign(new Error("OPENAI API KEY em falta"), { status: 400 });
      const model = mode === "quality" ? cfg.openai.modelQuality : cfg.openai.modelFast;
      return await openaiCreateResponse({
        apiKey: cfg.openai.apiKey,
        model,
        model,
        instructions,
        input,
        files, // Pass files to OpenAI (will handle images if supported)
        history, // NEW: Pass history
        max_output_tokens,
        temperature,
      });
    }

    if (selectedProvider === "gemini") {
      if (!cfg.gemini.apiKey) throw Object.assign(new Error("GEMINI API KEY em falta"), { status: 400 });
      const model = cfg.gemini.model || "gemini-2.0-flash";
      return await geminiCreateResponse({
        apiKey: cfg.gemini.apiKey,
        model,
        instructions,
        input,
        files,
        history, // NEW: Pass history
        max_output_tokens,
        temperature,
      });
    }
  } catch (err) {
    // --- FALLBACK LOGIC ---
    // If Gemini fails and we haven't tried OpenAI yet, fallback to OpenAI
    if (!isFallback && selectedProvider === "gemini" && cfg.openai.apiKey) {
      console.warn(`[ai] Gemini failed (${err.message}). Falling back to OpenAI with text extraction...`);

      let extraContext = "";
      for (const f of files) {
        if (f.type === "application/pdf" && f.content) {
          const buffer = Buffer.from(f.content, 'base64');
          const text = await extractTextFromPdfBuffer(buffer);
          if (text) {
            extraContext += `\n--- CONTEÚDO EXTRAÍDO DO PDF (${f.name}) ---\n${text}\n---\n`;
          }
        }
      }

      return await aiCreateText({
        mode,
        instructions: `[NOTA: O motor de análise visual falhou. O texto abaixo foi extraído manualmente do PDF.]\n\n${instructions}${extraContext}`,
        input,
        files: [],
        history, // NEW: Propagate history to fallback
        max_output_tokens,
        temperature,
        isFallback: true,
      });
    }
    throw err;
  }

  throw Object.assign(new Error(`AI_PROVIDER inválido: ${selectedProvider}`), { status: 400 });
}
