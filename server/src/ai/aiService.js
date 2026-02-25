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
    geminiModelFast: cfg.gemini.modelFast,
    geminiModelQuality: cfg.gemini.modelQuality,
  };
}

export async function aiCreateText(args) {
  const {
    mode = "fast",
    instructions,
    input,
    files = [],
    history = [],
    max_output_tokens = 256,
    temperature = 0.2,
    isFallback = false,
    customModels = {},
  } = args;

  const cfg = getAiConfig();
  if (!cfg.enabled) {
    throw Object.assign(new Error("AI desativado."), { status: 400 });
  }

  // --- STRATEGY: Prioritize OpenAI (Pedro prefers GPT empathy) ---
  // We only fallback to Gemini if OpenAI fails or has no key.
  const providers = [];

  // 1. Start with the preferred provider from config (usually openai)
  providers.push(cfg.provider);

  // 2. Add the alternative as failover
  const alternative = cfg.provider === "openai" ? "gemini" : "openai";
  providers.push(alternative);

  // Filter out providers without keys
  const availableProviders = providers.filter(p => {
    const key = p === "openai"
      ? (customModels.openaiApiKey || cfg.openai.apiKey)
      : (customModels.geminiApiKey || cfg.gemini.apiKey);
    return Boolean(key);
  });

  if (availableProviders.length === 0) {
    throw Object.assign(new Error("Nenhuma API Key disponível (OpenAI/Gemini)."), { status: 400 });
  }

  let lastError = null;

  for (const provider of availableProviders) {
    try {
      if (provider === "openai") {
        const apiKey = customModels.openaiApiKey || cfg.openai.apiKey;
        const model = customModels.openaiModelFast || (mode === "quality" ? cfg.openai.modelQuality : cfg.openai.modelFast);
        console.log(`[ai] Calling OpenAI (${model})...`);
        return await openaiCreateResponse({
          apiKey,
          model,
          instructions,
          input,
          files,
          history,
          max_output_tokens,
          temperature,
        });
      }

      if (provider === "gemini") {
        const apiKey = customModels.geminiApiKey || cfg.gemini.apiKey;
        const model = customModels.geminiModel || (mode === "quality" ? cfg.gemini.modelQuality : cfg.gemini.modelFast);
        console.log(`[ai] Calling Gemini (${model})...`);
        return await geminiCreateResponse({
          apiKey,
          model,
          instructions,
          input,
          files,
          history,
          max_output_tokens,
          temperature,
        });
      }
    } catch (err) {
      console.warn(`[ai] Provider ${provider} failed: ${err.message}`);
      lastError = err;

      // If we are on the first provider and it failed, and it was OpenAI with PDFs,
      // we might want to "pre-extract" text for the next provider if it's OpenAI too (unlikely)
      // but if the NEXT is Gemini, it handles PDFs natively.
      // If the CURRENT was Gemini and failed, and the NEXT is OpenAI, we MUST extract text.
      if (provider === "gemini" && availableProviders.indexOf(provider) < availableProviders.length - 1) {
        const nextProvider = availableProviders[availableProviders.indexOf(provider) + 1];
        if (nextProvider === "openai") {
          console.log("[ai] Falling back from Gemini to OpenAI, extracting PDF text...");
          let extraContext = "";
          for (const f of files) {
            if (f.type === "application/pdf" && f.content) {
              const buffer = Buffer.from(f.content, 'base64');
              const text = await extractTextFromPdfBuffer(buffer);
              if (text) extraContext += `\n--- CONTEÚDO EXTRAÍDO DO PDF (${f.name}) ---\n${text}\n---\n`;
            }
          }
          // Update instructions for the next attempt
          args.instructions = `[NOTA: O motor fallback foi ativado.]\n\n${instructions}${extraContext}`;
          args.files = []; // Clear files since OpenAI handles text now
        }
      }
    }
  }

  throw lastError || new Error("Falha ao contactar serviços de IA.");
}

export async function aiSelftest(customModels = {}) {
  const cfg = getAiConfig();
  const status = {
    ok: false,
    openai: { ok: false, error: null },
    gemini: { ok: false, error: null },
  };

  const checkOpenAI = async () => {
    const key = customModels.openaiApiKey || cfg.openai.apiKey;
    if (!key) return { ok: false, error: "Sem API Key" };
    try {
      await openaiCreateResponse({
        apiKey: key,
        model: customModels.openaiModelFast || cfg.openai.modelFast,
        instructions: "Responde apenas OK",
        input: "ping",
        max_output_tokens: 5,
        temperature: 0,
      });
      return { ok: true };
    } catch (e) {
      console.warn("[ai] OpenAI selftest failed:", e.message);
      return { ok: false, error: e.message };
    }
  };

  const checkGemini = async () => {
    const key = customModels.geminiApiKey || cfg.gemini.apiKey;
    if (!key) return { ok: false, error: "Sem API Key" };
    try {
      await geminiCreateResponse({
        apiKey: key,
        model: customModels.geminiModel || cfg.gemini.modelFast,
        instructions: "Responde apenas OK",
        input: "ping",
        max_output_tokens: 5,
        temperature: 0,
      });
      return { ok: true };
    } catch (e) {
      console.warn("[ai] Gemini selftest failed:", e.message);
      return { ok: false, error: e.message };
    }
  };

  const [oa, ge] = await Promise.all([checkOpenAI(), checkGemini()]);
  status.openai = oa;
  status.gemini = ge;
  status.ok = oa.ok || ge.ok;

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
