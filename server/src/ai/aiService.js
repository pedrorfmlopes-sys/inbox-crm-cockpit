// server/src/ai/aiService.js
import { getAiConfig } from "./aiConfig.js";
import { openaiCreateResponse } from "./openaiProvider.js";

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
  max_output_tokens = 256,
  temperature = 0.2,
}) {
  const cfg = getAiConfig();

  if (!cfg.enabled) {
    const err = new Error(
      "AI desativado. Define AI_ENABLED=1 e a respetiva API key no .env do server."
    );
    err.status = 400;
    throw err;
  }

  if (cfg.provider === "openai") {
    const model = mode === "quality" ? cfg.openai.modelQuality : cfg.openai.modelFast;
    return await openaiCreateResponse({
      apiKey: cfg.openai.apiKey,
      model,
      instructions,
      input,
      max_output_tokens,
      temperature,
    });
  }

  if (cfg.provider === "gemini") {
    const err = new Error("Provider Gemini ainda não implementado neste projeto (stub).");
    err.status = 501;
    throw err;
  }

  const err = new Error(`AI_PROVIDER inválido: ${cfg.provider}`);
  err.status = 400;
  throw err;
}
