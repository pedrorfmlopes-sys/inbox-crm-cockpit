import { Router } from "express";
import * as learningStore from "../learningStore.js";
import { aiCreateText } from "../ai/aiService.js";

export function createLearningRouter() {
    const router = Router();

    /**
     * Log a new interaction (called when user clicks "Insert")
     */
    router.post("/log", async (req, res) => {
        try {
            const log = req.body;
            if (!log.userResponse) {
                return res.status(400).json({ ok: false, message: "Missing userResponse" });
            }

            await learningStore.logInteraction(log);
            return res.json({ ok: true });
        } catch (e) {
            console.error("[learningRouter] Log error:", e);
            return res.status(500).json({ ok: false, error: e.message });
        }
    });

    /**
     * Get the current distilled profile (style & habits)
     */
    router.get("/profile", async (req, res) => {
        try {
            const userId = req.query.userId || "global";
            const profile = await learningStore.getStyleProfile(userId);
            return res.json({ ok: true, profile });
        } catch (e) {
            console.error("[learningRouter] Profile error:", e);
            return res.status(500).json({ ok: false, error: e.message });
        }
    });

    /**
     * Distill style and habits from logs
     */
    router.post("/distill", async (req, res) => {
        try {
            const logs = await learningStore.getLogs(30);
            if (logs.length < 3) {
                return res.json({ ok: false, message: "Necessário pelo menos 3 interações para destilar estilo." });
            }

            const logText = logs.map(l => (
                `ORIGINAL (${l.originalSubject}):\n${l.originalBody}\n` +
                `RESPOSTA DO PEDRO:\n${l.userResponse}\n` +
                `---`
            )).join("\n\n");

            const prompt = `
És um analista de comunicação especializado em espelhar estilos de escrita.
Analisa as seguintes interações reais do Pedro e extrai:
1. PERFIL DE ESTILO (StyleProfile): Descreve o tom, o vocabulário recorrente, a estrutura das frases e como ele saúda/fecha os emails.
2. HÁBITOS IDENTIFICADOS (HabitsData): Identifica padrões de fluxo (ex: "Sempre que recebe email de X, reencaminha para Y com o pedido Z").

INTERAÇÕES:
${logText}

Responde EXCLUSIVAMENTE em formato JSON:
{
  "styleData": "descrição detalhada do estilo",
  "habitsData": "descrição dos hábitos e automações sugeridas"
}
`;

            const result = await aiCreateText({
                mode: "quality",
                instructions: prompt,
                input: "Analisa o estilo e hábitos do Pedro.",
                max_output_tokens: 1000,
                temperature: 0.2
            });

            let distilled = { styleData: "", habitsData: "" };
            try {
                const jsonStr = result.text.substring(result.text.indexOf("{"), result.text.lastIndexOf("}") + 1);
                distilled = JSON.parse(jsonStr);
            } catch (e) {
                console.warn("[learningRouter] JSON Parse error during distillation:", result.text);
                return res.status(500).json({ ok: false, message: "Erro ao processar análise da IA." });
            }

            await learningStore.updateStyleProfile("global", distilled);

            return res.json({ ok: true, profile: distilled });
        } catch (e) {
            console.error("[learningRouter] Distill error:", e);
            return res.status(500).json({ ok: false, error: e.message });
        }
    });

    return router;
}
