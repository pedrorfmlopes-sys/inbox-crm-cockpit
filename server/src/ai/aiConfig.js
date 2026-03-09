// server/src/ai/aiConfig.js
export function getAiConfig() {
  const provider = (process.env.AI_PROVIDER || "openai").toLowerCase();
  const geminiModelFast = process.env.GEMINI_MODEL_FAST || "gemini-1.5-flash";
  const geminiModelQuality = process.env.GEMINI_MODEL_QUALITY || "gemini-1.5-pro";

  return {
    enabled: String(process.env.AI_ENABLED || "").trim() === "1",
    provider, // Primary/default provider
    openai: {
      apiKey: process.env.OPENAI_API_KEY || "",
      modelFast: process.env.OPENAI_MODEL_FAST || "gpt-4o-mini",
      modelQuality:
        process.env.OPENAI_MODEL_QUALITY || "gpt-4o",
    },
    gemini: {
      apiKey: process.env.GEMINI_API_KEY || "",
      modelFast: geminiModelFast,
      modelQuality: geminiModelQuality,
    },
  };
}
