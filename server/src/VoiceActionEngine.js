import { geminiCreateResponse } from "./ai/geminiProvider.js";
import { getAiConfig } from "./ai/aiConfig.js";

/**
 * VoiceActionEngine.js
 * Parses complex voice intents and orchestrates chained workflows.
 */

const MODEL_PRO = "gemini-1.5-pro";

export async function processVoiceCommand(commandText, context) {
    const cfg = getAiConfig();
    if (!cfg.gemini.apiKey) throw new Error("Missing Gemini API Key");

    const instructions = `
You are the Orchestrator for Cockpit V2. Parse the user's intent and return a sequence of actions.
The available actions are:
1. EXTRACT_ANCHORS: Extract project fingerprint from email body.
2. SCAN_PROTECTION: Check if project is protected in the local MOAT.
3. DRAFT_REJECTION: Generate a diplomatic rejection draft.
4. GENERATE_BRIEFING: Generate a 30-second briefing summary.
5. GENERATE_VCARD: Create a vCard for the contact.

User Command: "${commandText}"

Return the sequence as a valid JSON array of action keys.
Example: ["EXTRACT_ANCHORS", "SCAN_PROTECTION", "DRAFT_REJECTION"]
`;

    const result = await geminiCreateResponse({
        apiKey: cfg.gemini.apiKey,
        model: MODEL_PRO,
        instructions,
        input: `Context: ${JSON.stringify(context)}`,
        temperature: 0,
    });

    try {
        const cleaned = result.text.replace(/```json/g, "").replace(/```/g, "").trim();
        const actions = JSON.parse(cleaned);
        return { ok: true, actions };
    } catch (e) {
        console.error("[VoiceActionEngine] Failed to parse intent:", result.text);
        return { ok: false, error: "Cloud not understand intent" };
    }
}
