import { getAiConfig } from "./aiConfig.js";
import { openaiCreateResponse, openaiListModels } from "./openaiProvider.js";
import { geminiCreateResponse, geminiListModels } from "./geminiProvider.js";
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
  customModels = {}, // NEW: Allow overriding models from client
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
      const apiKey = customModels.openaiApiKey || cfg.openai.apiKey;
      if (!apiKey) throw Object.assign(new Error("OPENAI_API_KEY em falta"), { status: 400 });
      const model = customModels.openaiModelFast || (mode === "quality" ? cfg.openai.modelQuality : cfg.openai.modelFast);
      return await openaiCreateResponse({
        apiKey,
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
      const apiKey = customModels.geminiApiKey || cfg.gemini.apiKey;
      if (!apiKey) throw Object.assign(new Error("GEMINI_API_KEY em falta"), { status: 400 });
      const model = customModels.geminiModel || cfg.gemini.model || "gemini-1.5-flash";
      return await geminiCreateResponse({
        apiKey,
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
        customModels, // Propagate custom models
      });
    }
    throw err;
  }

  throw Object.assign(new Error(`AI_PROVIDER inválido: ${selectedProvider}`), { status: 400 });
}

export async function aiSelftest(customModels = {}) {
  const cfg = getAiConfig();
  const status = {
    ok: false,
    openai: false,
    gemini: false,
  };

  const checkOpenAI = async () => {
    const key = customModels.openaiApiKey || cfg.openai.apiKey;
    if (!key) return false;
    try {
      await openaiCreateResponse({
        apiKey: key,
        model: customModels.openaiModelFast || cfg.openai.modelFast,
        instructions: "Responde apenas OK",
        input: "ping",
        max_output_tokens: 5,
        temperature: 0,
      });
      return true;
    } catch (e) {
      console.warn("[ai] OpenAI selftest failed:", e.message);
      return false;
    }
  };

  const checkGemini = async () => {
    const key = customModels.geminiApiKey || cfg.gemini.apiKey;
    if (!key) return false;
    try {
      await geminiCreateResponse({
        apiKey: key,
        model: customModels.geminiModel || cfg.gemini.model || "gemini-1.5-flash",
        instructions: "Responde apenas OK",
        input: "ping",
        max_output_tokens: 5,
        temperature: 0,
      });
      return true;
    } catch (e) {
      console.warn("[ai] Gemini selftest failed:", e.message);
      return false;
    }
  };

  const [oa, ge] = await Promise.all([checkOpenAI(), checkGemini()]);
  status.openai = oa;
  status.gemini = ge;
  status.ok = oa || ge; // Overall OK if at least one works

  return status;
}

export async function listAvailableModels() {
  const cfg = getAiConfig();
  const results = {
    openai: [],
    gemini: [],
  };

  if (cfg.openai.apiKey) {
    try {
      results.openai = await openaiListModels(cfg.openai.apiKey);
    } catch (e) {
      console.warn("[ai] Failed to list OpenAI models:", e.message);
    }
  }

  if (cfg.gemini.apiKey) {
    try {
      results.gemini = await geminiListModels(cfg.gemini.apiKey);
    } catch (e) {
      console.warn("[ai] Failed to list Gemini models:", e.message);
    }
  }

  return results;
}
