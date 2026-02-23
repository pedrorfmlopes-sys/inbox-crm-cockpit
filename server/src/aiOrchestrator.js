import { aiCreateText } from "./ai/aiService.js";
import { getAiConfig } from "./ai/aiConfig.js";

/**
 * Cockpit V2 AI Orchestrator
 * Implements Multi-model strategy: Flash for speed (Extraction), Pro for reasoning (Analysis).
 * Fallback to default provider (OpenAI) if Gemini fails or is not preferred.
 */

const extractionInstructions = `
You are an expert industrial sales assistant. Your task is to extract tactical "anchors" from an email to help link it to Odoo.
Extract the following in JSON format:
{
  "projectName": "Name of the project if mentioned",
  "refArticles": ["SKU1", "SKU2"],
  "stakeholders": ["Name 1", "Name 2"],
  "location": "City/Country of the project",
  "participants": [
    { "name": "Name", "email": "email@host", "role": "sender|to|cc", "companyHint": "extracted from signature or domain" }
  ]
}
Only return the JSON. No conversational filler.
`;

export async function extractAnchors(emailContext, customModels = {}) {
    const cfg = getAiConfig();
    const apiKey = customModels.geminiApiKey || customModels.openaiApiKey || cfg.gemini.apiKey || cfg.openai.apiKey;
    if (!apiKey) throw new Error("Missing AI API Key (Gemini or OpenAI). Please check Settings.");

    const result = await aiCreateText({
        mode: "fast",
        instructions: extractionInstructions,
        input: typeof emailContext === 'string' ? emailContext : JSON.stringify(emailContext),
        temperature: 0.1,
        max_output_tokens: 800,
        customModels,
    });

    try {
        // Clean markdown backticks if any
        const cleaned = result.text.replace(/```json/g, "").replace(/```/g, "").trim();
        return JSON.parse(cleaned);
    } catch (e) {
        console.error("[aiOrchestrator] Failed to parse anchors JSON:", result.text);
        return { projectName: "", refArticles: [], stakeholders: [], location: "", participants: [] };
    }
}

export async function generateExecutiveSummary(context, emailHistory = [], customModels = {}) {
    const cfg = getAiConfig();
    const apiKey = customModels.geminiApiKey || customModels.openaiApiKey || cfg.gemini.apiKey || cfg.openai.apiKey;
    if (!apiKey) throw new Error("Missing AI API Key (Gemini or OpenAI). Please check Settings.");

    const instructions = `
You are the "Second Brain" assistant for an industrial sales expert.
Generate a "30-Second Briefing" in 3 bullet points:
1. Last critical steps from Outlook history.
2. Key notes from Odoo Chatter.
3. Protection status (Free or Protected).

Be extremely concise (max 40 words total).
`;

    try {
        const result = await aiCreateText({
            mode: "quality",
            instructions,
            input: context,
            history: emailHistory,
            temperature: 0.3,
            max_output_tokens: 300,
            customModels,
        });
        return result.text;
    } catch (e) {
        console.error("[aiOrchestrator] Executive summary generation failed:", e.message);
        throw e;
    }
}
