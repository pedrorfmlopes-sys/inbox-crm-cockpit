/**
 * HistorySummaryService.ts
 * Client-side service to handle the "30-Second Briefing" synthesis.
 */

import { aiGenerateBriefing } from "../../api";

export interface BriefingContext {
    outlookHistory?: string;
    odooChatter?: string;
    protectionStatus?: string;
}

/**
 * Generates a 30-second briefing using multi-source synthesis.
 */
export async function get30SecondBriefing(ctx: BriefingContext, customModels?: any): Promise<string> {
    const combinedContext = `
OUTLOOK RECENT HISTORY:
${ctx.outlookHistory || "No recent history."}

ODOO CHATTER/NOTES:
${ctx.odooChatter || "No internal notes."}

PROTECTION MOAT STATUS:
${ctx.protectionStatus || "Unknown"}
`.trim();

    try {
        const res = await aiGenerateBriefing(combinedContext, [], customModels);
        if (res.ok) {
            return res.summary;
        }
        throw new Error("Briefing generation failed");
    } catch (e) {
        console.error("[HistorySummaryService] Error:", e);
        return "Erro ao gerar resumo. Tente novamente.";
    }
}
