// server/src/ai/aiConfig.js
export function getAiConfig() {
  const provider = (process.env.AI_PROVIDER || "openai").toLowerCase();

  return {
    enabled:
      String(process.env.AI_ENABLED || "").trim() === "1" ||
      Boolean(process.env.OPENAI_API_KEY || process.env.GEMINI_API_KEY),
    provider,
    openai: {
      apiKey: process.env.OPENAI_API_KEY || "",
      modelFast: process.env.OPENAI_MODEL_FAST || "gpt-5",
      modelQuality:
        process.env.OPENAI_MODEL_QUALITY || process.env.OPENAI_MODEL_FAST || "gpt-5",
    },
    gemini: {
      apiKey: process.env.GEMINI_API_KEY || "",
      model: process.env.GEMINI_MODEL || "gemini-1.5-pro",
    },
  };
}
